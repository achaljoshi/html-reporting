'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { ATS, kpiSamples, readWb, sampleWb, templateWb, edit } = require('../helpers/app');
const K = ATS.kpi, D = ATS.date;

const SCORECARD = ['coverage', 'leakage', 'dde', 'aging', 'automation', 'milestones', 'environment', 'tsr', 'csat', 'resource', 'demand', 'commercial', 'raid'];
const EXTRA = ['execution', 'readiness'];
const RAGS = ['Green', 'Amber', 'Red', 'Grey'];
const TODAY = '2026-12-31'; // fixed "today" so the computation never depends on the real clock

const samples = kpiSamples().map((s) => ({ folder: s.folder, data: K.parse(readWb(s.file)) }));

test('there are five sample portfolios and each parses with weeks, defects, TSRs and milestones', () => {
  assert.equal(samples.length, 5);
  samples.forEach(({ folder, data }) => {
    assert.equal(data.kind, 'kpi');
    assert.equal(data.config.portfolio, folder, 'Config > Portfolio Name matches the folder');
    assert.ok(data.snapshots.length >= 10, folder + ' weekly rows');
    assert.ok(data.defects.length > 0 && data.tsrs.length > 0 && data.milestones.length > 0, folder);
    assert.ok(data.holidays.length > 10, folder + ' bank holidays');
    assert.deepEqual(data.warnings, [], folder + ' sample data should parse without warnings');
    data.snapshots.forEach((s, i) => { if (i) assert.ok(s.weekEnding > data.snapshots[i - 1].weekEnding, 'snapshots ascend'); });
    assert.ok(data.snapshots.every((s) => D.dow(s.weekEnding) === 5), folder + ': all week endings are Fridays');
  });
});

test('computeAll: every sample portfolio yields the 13 scorecard KPIs + 2 extras with valid RAG values', () => {
  samples.forEach(({ folder, data }) => {
    const res = K.computeAll(data, K.monthPeriod(K.months(data).slice(-2)[0]), { today: TODAY });
    assert.equal(res.items.length, 15, folder);
    assert.deepEqual(res.scorecard.map((x) => x.key), SCORECARD, folder + ' scorecard order');
    assert.deepEqual(res.items.filter((x) => K.metaByKey[x.key].extra).map((x) => x.key), EXTRA);
    res.items.forEach((it) => {
      assert.ok(RAGS.includes(it.rag), folder + '/' + it.key + ' rag ' + it.rag);
      assert.ok(RAGS.includes(it.autoRag), folder + '/' + it.key + ' autoRag ' + it.autoRag);
      assert.equal(typeof it.position, 'string');
      assert.ok(!/NaN|undefined|Infinity|\[object/.test(it.position + ' ' + it.exec + ' ' + it.highlights.join(' ') + ' ' + it.concern.join(' ')), folder + '/' + it.key + ' text has a bad token');
    });
    assert.ok(['Green', 'Amber', 'Red'].includes(res.overall.rag), folder + ' overall');
    assert.equal(res.overall.computed, ATS.rag.worst(res.scorecard.map((x) => x.rag)), 'computed overall = worst scorecard RAG');
  });
});

test('computeAll never throws and never yields NaN text for any week or month of any sample portfolio', () => {
  samples.forEach(({ folder, data }) => {
    const periods = K.weeks(data).map(K.weekPeriod).concat(K.months(data).map(K.monthPeriod));
    periods.forEach((p) => {
      const res = K.computeAll(data, p, { today: TODAY });
      res.items.forEach((it) => {
        assert.notEqual(it.position, 'Error computing', folder + ' ' + p.key + ' ' + it.key);
        assert.ok(!/NaN|undefined/.test(it.position), folder + ' ' + p.key + ' ' + it.key + ': ' + it.position);
        if (it.metric && it.metric.value != null) assert.ok(Number.isFinite(it.metric.value), folder + ' ' + p.key + ' ' + it.key + ' metric');
      });
    });
  });
});

test('month roll-up: flow measures (test days, downtime) are summed over the weeks that end in the month', () => {
  samples.forEach(({ folder, data }) => {
    const mk = '2026-09', p = K.monthPeriod(mk);
    const weeks = data.snapshots.filter((s) => s.weekEnding >= p.start && s.weekEnding <= p.end);
    assert.ok(weeks.length >= 3, folder + ' has weeks in Sep 2026');
    const days = weeks.reduce((a, s) => a + s.testDays, 0), down = weeks.reduce((a, s) => a + s.downtime, 0);
    const env = K.computeAll(data, p, { today: TODAY }).byKey.environment;
    assert.equal(env.detail.days, days, folder + ' summed test days');
    assert.equal(env.detail.down, down, folder + ' summed downtime');
    assert.equal(env.detail.req, days * data.th.hoursPerTestDay);
    assert.ok(Math.abs(env.detail.avail - (1 - down / (days * data.th.hoursPerTestDay))) < 1e-12);
    // each single week of the month, summed, equals the month
    const perWeek = weeks.map((s) => K.computeAll(data, K.weekPeriod(s.weekEnding), { today: TODAY, withPrev: false }).byKey.environment.detail);
    assert.equal(perWeek.reduce((a, d) => a + d.down, 0), down);
    assert.equal(perWeek.reduce((a, d) => a + d.days, 0), days);
  });
});

test('month roll-up: snapshot measures take the latest week of the month', () => {
  samples.forEach(({ folder, data }) => {
    const p = K.monthPeriod('2026-09');
    const last = data.snapshots.filter((s) => s.weekEnding <= p.end).pop();
    const res = K.computeAll(data, p, { today: TODAY });
    assert.equal(res.overall.snapWeek, last.weekEnding, folder);
    assert.equal(res.byKey.coverage.detail.snapWeek, last.weekEnding);
    const asWeek = K.computeAll(data, K.weekPeriod(last.weekEnding), { today: TODAY });
    ['coverage', 'automation', 'resource', 'commercial', 'execution'].forEach((k) => {
      assert.equal(res.byKey[k].position, asWeek.byKey[k].position, folder + '/' + k + ': month == its last week');
      assert.equal(res.byKey[k].rag, asWeek.byKey[k].rag, folder + '/' + k);
    });
  });
});

test('"as of" semantics: a past month is not changed by data added later', () => {
  const { data } = samples.find((s) => s.folder === 'PAYE');
  const aug = K.computeAll(data, K.monthPeriod('2026-08'), { today: TODAY });
  const trimmed = JSON.parse(JSON.stringify(data));
  delete trimmed._hols; // cached Set, not JSON-cloneable
  trimmed.snapshots = trimmed.snapshots.filter((s) => s.weekEnding <= '2026-08-31');
  trimmed.defects = trimmed.defects.filter((d) => d.raised <= '2026-08-31');
  trimmed.topicProgress = trimmed.topicProgress.filter((t) => t.weekEnding <= '2026-08-31');
  const augTrim = K.computeAll(trimmed, K.monthPeriod('2026-08'), { today: TODAY });
  ['coverage', 'environment', 'execution', 'dde'].forEach((k) => assert.equal(augTrim.byKey[k].position, aug.byKey[k].position, k));
});

test('computeAll: period metadata and previous-period link', () => {
  const { data } = samples[0];
  const res = K.computeAll(data, K.monthPeriod('2026-09'), { today: TODAY });
  assert.equal(res.period.start, '2026-09-01');
  assert.equal(res.period.end, '2026-09-30');
  assert.equal(res.prevPeriod.key, '2026-08');
  assert.equal(res.prev.period.key, '2026-08');
  assert.equal(res.items[0].prev, res.prev.byKey[res.items[0].key]);
  const wk = K.computeAll(data, K.weekPeriod('2026-09-25'), { today: TODAY });
  assert.equal(wk.period.start, '2026-09-19');
  assert.equal(wk.prevPeriod.key, '2026-09-18');
  assert.equal(K.months(data).join(), '2026-07,2026-08,2026-09,2026-10');
});

test('computeAll: asOf is capped at today so a future period never reads beyond it', () => {
  const { data } = samples[0];
  const res = K.computeAll(data, K.monthPeriod('2026-09'), { today: '2026-09-10' });
  assert.equal(res.asOf, '2026-09-10');
  assert.equal(K.computeAll(data, K.monthPeriod('2026-09'), { today: TODAY }).asOf, '2026-09-30');
});

test('RAG overrides from Excel win over computed colours and keep both visible', () => {
  const wb = sampleWb('Self Assessment', 'ATS_Weekly_Data.xlsx');
  edit(wb, 'Commentary', 'Week Ending', (aoa, c) => {
    c.append({ 'Week Ending': '2026-09-25', KPI: 'Environment Availability', 'RAG Override': 'Red', 'Executive Commentary': 'Manual wording from Excel' });
    c.append({ 'Week Ending': '2026-09-25', KPI: 'Overall', 'RAG Override': 'Amber', 'Executive Commentary': 'Overall words' });
  });
  const data = K.parse(wb), res = K.computeAll(data, K.weekPeriod('2026-09-25'), { today: TODAY });
  const env = res.byKey.environment;
  assert.equal(env.rag, 'Red');
  assert.equal(env.ragOverridden, env.autoRag !== 'Red');
  assert.equal(env.exec, 'Manual wording from Excel');
  assert.equal(env.src.exec, 'excel');
  assert.equal(res.overall.rag, 'Amber');
  assert.equal(res.overall.declared, 'Amber');
  assert.equal(res.overall.summary, 'Overall words');
  assert.equal(res.byKey.coverage.src.exec, 'auto', 'KPIs without commentary stay auto-drafted');
});

test('computeAll: RAID KPI uses the RAID log when provided and is empty otherwise', () => {
  const kdata = samples.find((s) => s.folder === 'Self Assessment').data;
  const raid = ATS.raid.parse(sampleWb('Self Assessment', 'ATS_RAID_Log.xlsx'));
  const without = K.computeAll(kdata, K.monthPeriod('2026-09'), { today: TODAY }).byKey.raid;
  assert.equal(without.empty, true);
  assert.equal(without.rag, 'Grey');
  const withRaid = K.computeAll(kdata, K.monthPeriod('2026-09'), { today: TODAY, raid }).byKey.raid;
  assert.equal(withRaid.empty, false);
  assert.equal(withRaid.rag, 'Red', 'a Very High item is open');
  assert.equal(withRaid.metric.value, withRaid.detail.veryHigh + withRaid.detail.high);
});

test('the blank weekly template parses to zero rows and computes without errors', () => {
  const data = K.parse(templateWb('ATS_Weekly_Data_Template.xlsx'));
  assert.equal(data.snapshots.length + data.defects.length + data.tsrs.length + data.milestones.length, 0);
  assert.deepEqual(K.months(data), []);
  assert.deepEqual(K.weeks(data), []);
  const res = K.computeAll(data, K.monthPeriod('2026-09'), { today: TODAY });
  assert.equal(res.items.length, 15);
  res.items.forEach((it) => assert.notEqual(it.position, 'Error computing', it.key));
  assert.equal(res.byKey.coverage.empty, true);
  assert.equal(res.byKey.leakage.notInProd, true, 'no go-live date => Not in Production');
});

test('K.parse returns null for a non-KPI workbook', () => {
  assert.equal(K.parse(templateWb('ATS_RAID_Log_Template.xlsx')), null);
});

test('thresholds from Config override the defaults (and ratios accept % text)', () => {
  const wb = sampleWb('PAYE', 'ATS_Weekly_Data.xlsx');
  edit(wb, 'Config', 'Setting', (aoa, c) => {
    aoa.forEach((r) => { if (/^Defect Turnaround SLA/.test(r[0])) r[1] = 3; if (/^TSR SLA/.test(r[0])) r[1] = 7; if (/^Automation Target/i.test(r[0])) r[1] = '90%'; });
  });
  const th = K.parse(wb).th;
  assert.equal(th.turnSla, 3);
  assert.equal(th.tsrSla, 7);
  assert.equal(th.autoTarget, 0.9);
  assert.equal(th.hoursPerTestDay, 8, 'untouched settings keep their value');
});

test('defaultKey: weeks that are still in the future are ignored', () => {
  const past = ['2020-01-03', '2020-01-10', '2020-01-17'], future = ['2099-01-02', '2099-01-09'];
  assert.equal(K.defaultKey(past.concat(future), 'week', [past.concat(future).map((w) => ({ weekEnding: w }))]), '2020-01-17');
  assert.equal(K.defaultKey(past, 'week', [past.map((w) => ({ weekEnding: w }))]), '2020-01-17');
  assert.equal(K.defaultKey([], 'week', [[]]), null);
});

test('defaultKey: only-future data falls back to the latest period', () => {
  const future = ['2099-01-02', '2099-01-09'];
  assert.equal(K.defaultKey(future, 'week', [future.map((w) => ({ weekEnding: w }))]), '2099-01-09');
});

test('defaultKey: a month needs three started weeks, otherwise the previous month is shown', () => {
  const snaps = (ws) => [ws.map((w) => ({ weekEnding: w }))];
  const twoWeeksOfFeb = ['2020-01-03', '2020-01-10', '2020-01-17', '2020-01-24', '2020-01-31', '2020-02-07', '2020-02-14'];
  assert.equal(K.defaultKey(['2020-01', '2020-02'], 'month', snaps(twoWeeksOfFeb)), '2020-01');
  const threeWeeksOfFeb = twoWeeksOfFeb.concat(['2020-02-21']);
  assert.equal(K.defaultKey(['2020-01', '2020-02'], 'month', snaps(threeWeeksOfFeb)), '2020-02');
  const withFuture = twoWeeksOfFeb.concat(['2099-03-06', '2099-03-13', '2099-03-20']);
  assert.equal(K.defaultKey(['2020-01', '2020-02', '2099-03'], 'month', snaps(withFuture)), '2020-01', 'pre-filled future months are skipped');
  assert.equal(K.defaultKey(['2020-01'], 'month', snaps(['2020-01-10'])), '2020-01', 'a single month is always shown');
});

test('weeklySeries returns one computed result per week, ending at the requested week', () => {
  const { data } = samples[0];
  const s = K.weeklySeries(data, '2026-09-25', 4, { today: TODAY });
  assert.equal(s.length, 4);
  assert.equal(s[3].we, '2026-09-25');
  assert.deepEqual(s.map((x) => x.we), ['2026-09-04', '2026-09-11', '2026-09-18', '2026-09-25']);
  s.forEach((x) => assert.equal(x.res.items.length, 15));
});

test('delta: direction and goodness follow the metric\'s "better" direction', () => {
  const mk = (cur, prev, better, fmt) => ({ metric: { value: cur, better, fmt: fmt || 'pct' }, prev: { metric: { value: prev } } });
  const up = K.delta(mk(0.9, 0.8, 'up'));
  assert.equal(up.dir, 1); assert.equal(up.good, true); assert.equal(up.text, '▲ 10.0 pts');
  const downBad = K.delta(mk(0.7, 0.8, 'up'));
  assert.equal(downBad.dir, -1); assert.equal(downBad.good, false);
  assert.equal(K.delta(mk(3, 5, 'down', 'int')).good, true);
  assert.equal(K.delta(mk(5, 5, 'up')).text, 'no change');
  assert.equal(K.delta({ metric: { value: 1 } }), null, 'no previous period');
  assert.equal(K.delta(mk(null, 5, 'up')), null);
});
