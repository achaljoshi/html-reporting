'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { ATS, sampleWb, edit, lastFriday } = require('../helpers/app');
const K = ATS.kpi, D = ATS.date;

const base = () => sampleWb('Self Assessment', 'ATS_Weekly_Data.xlsx');
const msgs = (data, sheet) => K.quality(data).filter((w) => !sheet || w.sheet === sheet).map((w) => w.msg);
const has = (list, re) => list.some((m) => re.test(m));
// keep only the first n weekly rows so the data is certain to be old (the sample's own dates are in 2026)
const truncate = (wb, n) => edit(wb, 'Weekly_Snapshot', 'Week Ending', (aoa, c) => { c.dataRows().slice(n).forEach((x) => aoa[x.i].fill('')); });
// make the latest weekly row a recent Friday so the data is "fresh" whatever the real date is
const freshen = (wb) => edit(wb, 'Weekly_Snapshot', 'Week Ending', (aoa, c) => { const rows = c.dataRows(); aoa[rows[rows.length - 1].i][c.col('Week Ending')] = lastFriday(0); });

test('quality: a clean, fresh sample workbook has no warnings', () => {
  const data = K.parse(freshen(base()));
  assert.deepEqual(K.quality(data), []);
});

test('quality: text in a number cell is reported and the value ignored', () => {
  const wb = base();
  edit(wb, 'Weekly_Snapshot', 'Week Ending', (aoa, c) => { const r = c.dataRows()[2]; aoa[r.i][c.col('Downtime Hours')] = 'about four'; });
  const data = K.parse(wb);
  const w = data.warnings.filter((x) => x.sheet === 'Weekly_Snapshot');
  assert.equal(w.length, 1);
  assert.match(w[0].msg, /"about four" in Downtime Hours is not a number/);
  assert.match(w[0].msg, /^W\/E 2026-07-24/);
  assert.equal(data.snapshots[2].downtime, null);
  assert.ok(has(msgs(data, 'Weekly_Snapshot'), /not a number/), 'also surfaced by K.quality');
});

test('quality: a weekly row with an unreadable date is dropped and reported', () => {
  const wb = base();
  const before = K.parse(base()).snapshots.length;
  edit(wb, 'Weekly_Snapshot', 'Week Ending', (aoa, c) => { aoa[c.dataRows()[3].i][c.col('Week Ending')] = '2026-25-09'; });
  const data = K.parse(wb);
  assert.equal(data.snapshots.length, before - 1);
  assert.ok(has(msgs(data, 'Weekly_Snapshot'), /"2026-25-09" is not a valid date — this row is not counted/));
});

test('quality: an empty Week Ending on a row with data is reported', () => {
  const wb = base();
  edit(wb, 'Weekly_Snapshot', 'Week Ending', (aoa, c) => { aoa[c.dataRows()[1].i][c.col('Week Ending')] = ''; });
  assert.ok(has(msgs(K.parse(wb), 'Weekly_Snapshot'), /Week Ending is empty/));
});

test('quality: defects / TSRs / milestones with unreadable dates or no ID are reported, not silently lost', () => {
  const wb = base();
  edit(wb, 'Defect_Log', 'Defect ID', (aoa, c) => { aoa[c.dataRows()[0].i][c.col('Raised Date')] = 'next Tuesday'; aoa[c.dataRows()[1].i][c.col('Defect ID')] = ''; });
  edit(wb, 'TSR_Log', 'TSR Ref', (aoa, c) => { aoa[c.dataRows()[0].i][c.col('Received Date')] = '31/02/2026'; });
  edit(wb, 'Milestones', 'Milestone', (aoa, c) => { aoa[c.dataRows()[0].i][c.col('Due Date')] = ''; });
  const clean = K.parse(base());
  const data = K.parse(wb);
  assert.equal(data.defects.length, clean.defects.length - 2);
  assert.equal(data.tsrs.length, clean.tsrs.length - 1);
  assert.equal(data.milestones.length, clean.milestones.length - 1);
  assert.ok(has(msgs(data, 'Defect_Log'), /Raised Date "next Tuesday" is not a valid date/));
  assert.ok(has(msgs(data, 'Defect_Log'), /Defect ID is empty/));
  assert.ok(has(msgs(data, 'TSR_Log'), /Received Date "31\/02\/2026" is not a valid date/));
  assert.ok(has(msgs(data, 'Milestones'), /Due Date is empty/));
});

test('quality: UK day-first dates in cells are understood, not flagged', () => {
  const wb = base();
  edit(wb, 'Defect_Log', 'Defect ID', (aoa, c) => { aoa[c.dataRows()[0].i][c.col('Raised Date')] = '25/09/2026'; aoa[c.dataRows()[1].i][c.col('Raised Date')] = 'Sep 25, 2026'; });
  const data = K.parse(wb);
  assert.deepEqual(data.warnings, []);
  assert.equal(data.defects.find((d) => d.raised === '2026-09-25') !== undefined, true);
});

test('quality: future (pre-filled) weeks are flagged and still counted', () => {
  const wb = freshen(base());
  const future = '2099-01-02'; // a Friday
  assert.equal(D.dow(future), 5);
  edit(wb, 'Weekly_Snapshot', 'Week Ending', (aoa, c) => { c.append({ 'Week Ending': future }); });
  const data = K.parse(wb);
  assert.ok(data.snapshots.some((s) => s.weekEnding === future));
  assert.ok(has(msgs(data, 'Weekly_Snapshot'), /W\/E 2099-01-02 is in the future — it looks pre-filled/));
});

test('quality: stale data (latest week more than 9 days old) is flagged, fresh data is not', () => {
  const stale = K.parse(truncate(base(), 4));
  const w = msgs(stale, 'Weekly_Snapshot').filter((m) => /days old/.test(m));
  assert.equal(w.length, 1);
  assert.match(w[0], /Latest weekly snapshot is \d+ days old \(W\/E 2026-07-31\) — add this week's row/);
  assert.ok(!has(msgs(K.parse(freshen(base()))), /days old/));
});

test('quality: a week ending that is not a Friday is flagged', () => {
  const wb = freshen(base());
  edit(wb, 'Weekly_Snapshot', 'Week Ending', (aoa, c) => { aoa[c.dataRows()[1].i][c.col('Week Ending')] = '2026-07-15'; }); // a Wednesday
  assert.ok(has(msgs(K.parse(wb), 'Weekly_Snapshot'), /W\/E 2026-07-15 is not a Friday/));
});

test('quality: coverage / automation numbers that cannot be right are flagged', () => {
  const wb = freshen(base());
  edit(wb, 'Weekly_Snapshot', 'Week Ending', (aoa, c) => {
    const r = c.dataRows()[2].i;
    aoa[r][c.col('High Risk – Covered Reqs')] = 500; // more covered than requirements (72)
    aoa[r][c.col('Automated Test Data Setups')] = 9999;
  });
  const m = msgs(K.parse(wb), 'Weekly_Snapshot');
  assert.ok(has(m, /high risk covered requirements exceed total requirements/));
  assert.ok(has(m, /automated test-data setups exceed automatable/));
});

test('quality: defect status / date contradictions are flagged', () => {
  const wb = freshen(base());
  edit(wb, 'Defect_Log', 'Defect ID', (aoa, c) => {
    const rows = c.dataRows();
    aoa[rows[0].i][c.col('Status')] = 'Closed'; aoa[rows[0].i][c.col('Resolved Date')] = '';           // closed, no resolved date
    aoa[rows[1].i][c.col('Raised Date')] = '2026-09-10'; aoa[rows[1].i][c.col('Resolved Date')] = '2026-09-01'; // resolved before raised
    aoa[rows[2].i][c.col('Status')] = 'Open'; aoa[rows[2].i][c.col('Resolved Date')] = '2026-09-02';           // open but resolved
  });
  const data = K.parse(wb), m = msgs(data, 'Defect_Log');
  assert.ok(has(m, new RegExp(data.defects[0].id + ' is Closed but has no Resolved Date')));
  assert.ok(has(m, new RegExp(data.defects[1].id + ': Resolved Date is before Raised Date')));
  assert.ok(has(m, new RegExp(data.defects[2].id + ' has a Resolved Date but status is Open')));
});

test('quality: TSR returned before received is flagged', () => {
  const wb = freshen(base());
  edit(wb, 'TSR_Log', 'TSR Ref', (aoa, c) => { const r = c.dataRows()[1].i; aoa[r][c.col('Received Date')] = '2026-06-10'; aoa[r][c.col('Returned Date')] = '2026-06-01'; });
  assert.ok(has(msgs(K.parse(wb), 'TSR_Log'), /Returned Date is before Received Date/));
});

test('quality: the "today" argument is accepted without changing results for clean data', () => {
  const data = K.parse(freshen(base()));
  assert.deepEqual(K.quality(data, '2030-01-01'), K.quality(data));
});

test('parse: classification of defects (priority, status, found-in) is normalised', () => {
  const wb = base();
  edit(wb, 'Defect_Log', 'Defect ID', (aoa, c) => {
    const rows = c.dataRows();
    aoa[rows[0].i][c.col('Priority')] = 'critical'; aoa[rows[0].i][c.col('Status')] = 'rejected - duplicate';
    aoa[rows[1].i][c.col('Priority')] = 'medium'; aoa[rows[1].i][c.col('Status')] = 'in progress'; aoa[rows[1].i][c.col('Found In')] = 'Production';
    aoa[rows[2].i][c.col('Priority')] = 'P3'; aoa[rows[2].i][c.col('Status')] = 're-open';
  });
  const d = K.parse(wb).defects;
  assert.equal(d[0].priority, 'P1 High'); assert.equal(d[0].status, 'Rejected');
  assert.equal(d[1].priority, 'P2 Medium'); assert.equal(d[1].status, 'In Progress'); assert.equal(d[1].foundIn, 'Production');
  assert.equal(d[2].priority, 'P3 Low'); assert.equal(d[2].status, 'Re-Open');
});

test('rows whose Notes start with EXAMPLE are ignored', () => {
  const wb = base();
  const n = K.parse(base()).defects.length;
  edit(wb, 'Defect_Log', 'Defect ID', (aoa, c) => { c.append({ 'Defect ID': 'EX-1', Summary: 'example', Priority: 'P1', 'Raised Date': '2026-09-01', Status: 'Open', Notes: 'EXAMPLE row — delete me' }); });
  assert.equal(K.parse(wb).defects.length, n);
});
