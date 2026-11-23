'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { ATS, sampleWb, templateWb, edit, dropSheet } = require('../helpers/app');
const RA = ATS.raid;

const sample = () => sampleWb('Self Assessment', 'ATS_RAID_Log.xlsx');
const parsed = RA.parse(sample());
const qmsgs = (raid) => RA.quality(raid).map((w) => w.sheet + ': ' + w.msg);

test('parse: sample RAID log has all five registers and its summary-sheet details', () => {
  assert.equal(parsed.kind, 'raid');
  assert.deepEqual([parsed.risks.length, parsed.issues.length, parsed.deps.length, parsed.assumptions.length, parsed.decisions.length], [14, 5, 8, 6, 6]);
  assert.equal(parsed.info.project, 'Self Assessment');
  assert.equal(parsed.info.start, '2026-07-01');
  assert.equal(parsed.info.end, '2027-06-30');
  assert.match(RA.describe(parsed), /14 risks · 5 issues · 8 dependencies · 6 assumptions · 6 decisions/);
});

test('parse: risk score is likelihood x impact and maps to a rating band', () => {
  parsed.risks.forEach((r) => {
    assert.ok(r.id && r.title, 'risk has id and title');
    assert.ok(r.l >= 1 && r.l <= 5 && r.i >= 1 && r.i <= 5, r.id + ' l/i in 1..5');
    assert.equal(r.score, r.l * r.i, r.id);
    assert.equal(r.rating, RA.bandOf(r.score), r.id);
  });
  const r1 = parsed.risks.find((r) => r.id === 'R-001');
  assert.deepEqual([r1.l, r1.i, r1.score, r1.rating], [4, 4, 16, 'High']);
  assert.equal(parsed.risks.find((r) => r.id === 'R-003').rating, 'Very High');
});

test('bandOf: boundaries of the five bands', () => {
  const cases = [[1, 'Very Low'], [4, 'Very Low'], [5, 'Low'], [9, 'Low'], [10, 'Medium'], [14, 'Medium'], [15, 'High'], [19, 'High'], [20, 'Very High'], [25, 'Very High'], [0, 'Very Low'], [null, null]];
  cases.forEach(([s, label]) => assert.equal(RA.bandOf(s), label, 'score ' + s));
});

test('parse: every item carries an ID and archived flags default to false', () => {
  [parsed.risks, parsed.issues, parsed.deps, parsed.assumptions, parsed.decisions].forEach((list) => list.forEach((x) => { assert.ok(x.id); assert.equal(typeof x.archived, 'boolean'); }));
});

test('summary: counts, bands and top lists are consistent with the open items', () => {
  const s = RA.summary(parsed, '2026-09-30', { start: '2026-09-01', end: '2026-09-30' });
  assert.equal(s.openRisks, parsed.risks.filter(RA.isOpenRisk).length);
  assert.equal(s.openIssues, parsed.issues.filter(RA.isOpenIssue).length);
  assert.equal(s.openDeps, parsed.deps.filter(RA.isOpenDep).length);
  assert.equal(s.riskBands.reduce((a, b) => a + b.n, 0), s.openRisks, 'risk bands add up to the open risks');
  assert.equal(s.issueBands.reduce((a, b) => a + b.n, 0), s.openIssues);
  assert.deepEqual(s.riskBands.map((b) => b.label), ['Very Low', 'Low', 'Medium', 'High', 'Very High']);
  assert.equal(s.veryHigh, s.riskBands[4].n + s.issueBands[4].n);
  assert.equal(s.high, s.riskBands[3].n + s.issueBands[3].n);
  assert.ok(s.veryHigh >= 1);
  assert.equal(s.topRisks.length, Math.min(6, s.openRisks));
  for (let i = 1; i < s.topRisks.length; i++) assert.ok(s.topRisks[i - 1].score >= s.topRisks[i].score, 'top risks sorted by score');
  assert.equal(Object.values(s.depByPriority).reduce((a, b) => a + b, 0), s.openDeps);
  assert.equal(s.decisionsTotal, parsed.decisions.filter((d) => !d.archived).length);
  assert.equal(s.decisionsPending, 1);
  assert.equal(s.pendingDecisions.length, s.decisionsPending);
});

test('summary: period counts only include items dated inside the period', () => {
  const sep = RA.summary(parsed, '2026-09-30', { start: '2026-09-01', end: '2026-09-30' });
  const none = RA.summary(parsed, '2020-01-31', { start: '2020-01-01', end: '2020-01-31' });
  assert.equal(none.newInPeriod, 0);
  assert.equal(none.closedInPeriod, 0);
  assert.equal(none.decisionsInPeriod, 0);
  assert.equal(sep.decisionsInPeriod, parsed.decisions.filter((d) => d.decided >= '2026-09-01' && d.decided <= '2026-09-30' && !d.archived).length);
  assert.ok(sep.decisionsInPeriod >= 3);
});

test('summary: dependencies due within 30 days of the as-of date, soonest first', () => {
  const s = RA.summary(parsed, '2026-10-01', { start: '2026-10-01', end: '2026-10-31' });
  s.depsDue.forEach((d) => assert.ok(d.required <= '2026-10-31'));
  for (let i = 1; i < s.depsDue.length; i++) assert.ok(s.depsDue[i - 1].required <= s.depsDue[i].required);
});

test('assumptionState: one definition of "unconfirmed" shared by every page (regression)', () => {
  const st = (status) => RA.assumptionState({ status });
  ['Unconfirmed', 'unconfirmed', 'Not yet confirmed', 'Open', '', 'TBC', 'Unvalidated'].forEach((s) => assert.equal(st(s), 'Unconfirmed', JSON.stringify(s)));
  ['Confirmed Correct', 'Confirmed', 'Validated', 'Correct'].forEach((s) => assert.equal(st(s), 'Confirmed Correct', s));
  ['Confirmed Incorrect', 'Incorrect', 'Proven incorrect'].forEach((s) => assert.equal(st(s), 'Confirmed Incorrect', s));
  ['Closed', 'Withdrawn', 'Superseded', 'Invalid'].forEach((s) => assert.equal(st(s), 'Closed', s));
});

test('summary.unconfirmedAssumptions agrees with assumptionState on the sample', () => {
  const s = RA.summary(parsed, '2026-09-30', { start: '2026-09-01', end: '2026-09-30' });
  const expected = parsed.assumptions.filter((a) => RA.assumptionState(a) === 'Unconfirmed').length;
  assert.equal(s.unconfirmedAssumptions, expected);
  assert.ok(expected > 0, 'the sample has unconfirmed assumptions');
});

test('open/closed helpers', () => {
  assert.equal(RA.isOpenRisk({ status: 'Open' }), true);
  assert.equal(RA.isOpenRisk({ status: 'Closed' }), false);
  assert.equal(RA.isOpenRisk({ status: 'Open', archived: true }), false);
  assert.equal(RA.isOpenIssue({ status: 'Resolved' }), false);
  assert.equal(RA.isOpenIssue({ status: 'Closed - fixed' }), false);
  assert.equal(RA.isOpenIssue({ status: 'In Progress' }), true);
  assert.equal(RA.isOpenDep({ status: 'Closed' }), false);
  assert.equal(RA.isPendingDecision({ status: 'Pending' }), true);
  assert.equal(RA.isPendingDecision({ status: 'Awaiting sign-off' }), true);
  assert.equal(RA.isPendingDecision({ status: 'Approved' }), false);
});

test('quality: the sample RAID log is clean', () => {
  assert.deepEqual(RA.quality(parsed), []);
  assert.deepEqual(RA.quality(null), []);
});

test('quality: duplicate IDs, issues without resolution dates, unrated items and undated decisions are flagged', () => {
  const wb = sample();
  edit(wb, 'Risks', 'ID', (aoa, c) => {
    const rows = c.dataRows();
    aoa[rows[1].i][c.col('ID')] = aoa[rows[0].i][c.col('ID')];       // duplicate ID
    aoa[rows[2].i][c.col('Likelihood')] = ''; aoa[rows[2].i][c.col('Impact')] = ''; // unrated open risk
    aoa[rows[2].i].forEach((v, i) => { if (/rating|score/i.test(String(aoa[c.hdr][i]))) aoa[rows[2].i][i] = ''; });
  });
  edit(wb, 'Issues', 'ID', (aoa, c) => { const r = c.dataRows()[0].i; aoa[r][c.col('Status')] = 'Resolved'; aoa[r][c.col('Actual Resolution Date')] = ''; });
  edit(wb, 'Decisions', 'ID', (aoa, c) => { const r = c.dataRows()[0].i; aoa[r][c.col('Status')] = 'Approved'; aoa[r][c.col('Date Decided')] = ''; });
  const bad = RA.parse(wb), m = qmsgs(bad);
  assert.ok(m.some((x) => /^Risks: ID R-001 appears on more than one row/.test(x)), m.join('\n'));
  assert.ok(m.some((x) => /^Risks: R-003 has no usable Likelihood/.test(x)), m.join('\n'));
  assert.ok(m.some((x) => /^Issues: .* is Resolved but has no Actual Resolution Date/.test(x)), m.join('\n'));
  assert.ok(m.some((x) => /^Decisions: .* is Approved but has no Date Decided/.test(x)), m.join('\n'));
});

test('Decisions sheet is optional (older RAID logs have none)', () => {
  const wb = dropSheet(sample(), 'Decisions');
  assert.equal(RA.detect(wb), true);
  const p = RA.parse(wb);
  assert.deepEqual(p.decisions, []);
  assert.equal(p.risks.length, 14);
  const s = RA.summary(p, '2026-09-30', { start: '2026-09-01', end: '2026-09-30' });
  assert.equal(s.decisionsTotal, 0);
  assert.equal(s.decisionsPending, 0);
  assert.deepEqual(RA.quality(p), []);
});

test('the Decisions sheet is found under alternative names', () => {
  const wb = sample();
  wb.SheetNames = wb.SheetNames.map((n) => (n === 'Decisions' ? 'Decision Log' : n));
  wb.Sheets['Decision Log'] = wb.Sheets['Decisions']; delete wb.Sheets['Decisions'];
  assert.equal(RA.parse(wb).decisions.length, 6);
});

test('archived rows are parsed but excluded from open counts', () => {
  const wb = sample();
  edit(wb, 'Risks', 'ID', (aoa, c) => { aoa[c.dataRows()[0].i][c.col('Archived')] = 'Yes'; });
  const p = RA.parse(wb), base = RA.summary(parsed, '2026-09-30', { start: '2026-09-01', end: '2026-09-30' }), s = RA.summary(p, '2026-09-30', { start: '2026-09-01', end: '2026-09-30' });
  assert.equal(p.risks.length, 14);
  assert.equal(s.openRisks, base.openRisks - 1);
});

test('blank RAID template parses to zero rows everywhere and still summarises', () => {
  const wb = templateWb('ATS_RAID_Log_Template.xlsx');
  assert.equal(RA.detect(wb), true);
  const p = RA.parse(wb);
  assert.deepEqual([p.risks.length, p.issues.length, p.deps.length, p.assumptions.length, p.decisions.length], [0, 0, 0, 0, 0]);
  assert.equal(RA.describe(p), '0 risks · 0 issues · 0 dependencies · 0 assumptions · 0 decisions');
  const s = RA.summary(p, '2026-09-30', { start: '2026-09-01', end: '2026-09-30' });
  assert.equal(s.openRisks + s.openIssues + s.openDeps + s.veryHigh + s.high, 0);
  assert.deepEqual(RA.quality(p), []);
});

test('parse returns null for a workbook that is not a RAID log', () => {
  assert.equal(RA.parse(templateWb('ATS_Weekly_Data_Template.xlsx')), null);
  assert.equal(RA.detect(templateWb('ATS_POAP_Template.xlsx')), false);
});

test('render helpers escape user-supplied text', () => {
  const evil = [{ id: '<b>1</b>', title: '<script>alert(1)</script>', rating: 'High', owner: '"x"', mitigation: "<img onerror='x'>" }];
  const html = RA.riskTable(evil);
  assert.ok(!/<script>|<img |<b>1<\/b>/.test(html), html);
  assert.ok(html.includes('&lt;script&gt;'));
  assert.match(RA.riskTable([]), /No open risks/);
});
