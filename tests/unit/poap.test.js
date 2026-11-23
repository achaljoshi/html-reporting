'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { ATS, sampleWb, templateWb, edit, dropSheet } = require('../helpers/app');
const P = ATS.poap;

const wbPoap = () => sampleWb('Self Assessment', 'ATS_POAP.xlsx');
const sample = P.parse(wbPoap());
const msgs = (data) => P.quality(data).map((w) => w.msg);
const has = (list, re) => list.some((m) => re.test(m));
// edit a row of POAP_Plan by bar ID
const editBar = (wb, id, fn) => edit(wb, 'POAP_Plan', 'ID', (aoa, c) => { const row = c.dataRows().find((x) => x.r[c.col('ID')] === id); assert.ok(row, 'bar ' + id + ' exists'); fn(aoa[row.i], c); });

test('parse: sample POAP has plan bars, milestones, P2P flows and config', () => {
  assert.ok(sample.plan.length >= 50);
  assert.ok(sample.milestones.length >= 5);
  assert.ok(sample.p2p.length >= 5);
  assert.equal(sample.config.statusDate, '2026-10-02');
  assert.equal(sample.config.start, '2026-06-01');
  assert.deepEqual(sample.config.pillars.map((p) => p.name), ['Pillar A – Returns', 'Pillar B – Payments']);
  assert.ok(sample.config.types.length >= 10);
  assert.ok(sample.config.types.every((t) => /^#[0-9a-f]{3,6}$/.test(t.color)), 'every activity type has a colour');
});

test('parse: bars have unique IDs, ISO dates and 0-1 percentages', () => {
  const ids = new Set();
  sample.plan.forEach((b) => {
    assert.ok(!ids.has(b.id), 'duplicate id ' + b.id); ids.add(b.id);
    ['start', 'end', 'baseStart', 'baseEnd'].forEach((k) => assert.ok(b[k] === '' || /^\d{4}-\d{2}-\d{2}$/.test(b[k]), b.id + ' ' + k + ' = ' + b[k]));
    if (b.start && b.end) assert.ok(b.end >= b.start, b.id + ' ends after it starts');
    if (b.pct != null) assert.ok(b.pct >= 0 && b.pct <= 1, b.id + ' pct');
    assert.ok(['', 'Green', 'Amber', 'Red'].includes(b.rag), b.id + ' rag');
  });
  const a1 = sample.plan.find((b) => b.id === 'A-001');
  assert.deepEqual([a1.pillar, a1.type, a1.start, a1.end, a1.pct, a1.status, a1.rag], ['Pillar A – Returns', 'Test Prep', '2026-06-15', '2026-07-31', 1, 'Complete', 'Green']);
});

test('parse: dependencies reference real bars, and P2P paths are split into system nodes', () => {
  const ids = new Set(sample.plan.map((b) => b.id));
  assert.ok(sample.plan.some((b) => b.dependsOn.length > 0));
  sample.plan.forEach((b) => b.dependsOn.forEach((d) => assert.ok(ids.has(d), b.id + ' depends on ' + d)));
  const f = sample.p2p[0];
  assert.deepEqual(f.pathNodes, ['Portal', 'API Gateway', 'Ledger']);
  ['r2xx', 'r3xx', 'r4xx', 'r5xx'].forEach((k) => assert.ok(['Pass', 'Fail', 'Blocked', 'N/A', 'Not Run'].includes(f[k]), k));
});

test('quality: the sample POAP is clean', () => {
  assert.deepEqual(P.quality(sample), []);
  assert.deepEqual(P.quality(null), []);
});

test('quality: duplicate bar IDs are renamed and reported', () => {
  const wb = wbPoap();
  editBar(wb, 'A-002', (row, c) => { row[c.col('ID')] = 'A-001'; });
  const d = P.parse(wb);
  assert.ok(d.plan.some((b) => b.id === 'A-001~2' && b.dupOf === 'A-001'), 'second use is renamed A-001~2');
  assert.equal(d.plan.filter((b) => b.id === 'A-001').length, 1);
  assert.ok(has(msgs(d), /ID A-001 is used on 2 rows/));
});

test('quality: a bar with no ID is reported (and gets a row-based ID)', () => {
  const wb = wbPoap();
  editBar(wb, 'A-003', (row, c) => { row[c.col('ID')] = ''; });
  const d = P.parse(wb);
  assert.ok(d.plan.some((b) => b.noId && /^ROW-\d+$/.test(b.id)));
  assert.ok(has(msgs(d), /has no ID, so it cannot be used as a dependency/));
});

test('quality: end date before start date', () => {
  const wb = wbPoap();
  editBar(wb, 'A-001', (row, c) => { row[c.col('Start')] = '2026-08-31'; row[c.col('End')] = '2026-08-01'; });
  const m = msgs(P.parse(wb));
  assert.ok(has(m, /^A-001 \(Test design & scripting\): End date is before Start date\.$/), m.join('\n'));
});

test('quality: unknown and self dependencies', () => {
  const wb = wbPoap();
  editBar(wb, 'A-002', (row, c) => { row[c.col('Depends On')] = 'A-001, ZZ-999'; });
  editBar(wb, 'A-003', (row, c) => { row[c.col('Depends On')] = 'A-003'; });
  const m = msgs(P.parse(wb));
  assert.ok(has(m, /A-002 .*depends on ZZ-999, which is not in the plan/), m.join('\n'));
  assert.ok(!has(m, /depends on A-001/), 'valid dependencies are not reported');
  assert.ok(has(m, /A-003 .*depends on itself/), m.join('\n'));
});

test('quality: missing dates are reported unless the bar is explicitly TBC', () => {
  const wb = wbPoap();
  editBar(wb, 'A-001', (row, c) => { row[c.col('End')] = ''; });
  editBar(wb, 'A-002', (row, c) => { row[c.col('Start')] = 'TBC'; row[c.col('End')] = 'TBC'; row[c.col('Status')] = 'Dates TBC'; });
  const d = P.parse(wb), m = msgs(d);
  assert.ok(has(m, /A-001 .*Start or End date is missing/), m.join('\n'));
  assert.ok(!has(m, /A-002 .*missing/), 'an explicit "Dates TBC" bar is fine');
  assert.equal(d.plan.find((b) => b.id === 'A-002').status, 'Dates TBC');
});

test('parse: impossible calendar dates are treated as missing, never rolled into another date (regression)', () => {
  const wb = wbPoap();
  editBar(wb, 'A-001', (row, c) => { row[c.col('Start')] = '2026-25-09'; });
  editBar(wb, 'A-002', (row, c) => { row[c.col('End')] = '31/02/2026'; });
  editBar(wb, 'A-003', (row, c) => { row[c.col('Start')] = '13/13/2026'; });
  const d = P.parse(wb), bar = (id) => d.plan.find((b) => b.id === id);
  assert.equal(bar('A-001').start, '');
  assert.equal(bar('A-002').end, '');
  assert.equal(bar('A-003').start, '');
  assert.ok(has(msgs(d), /A-001 .*Start or End date is missing/));
});

test('parse: assorted date formats in cells', () => {
  const wb = wbPoap();
  editBar(wb, 'A-001', (row, c) => { row[c.col('Start')] = '15/06/2026'; row[c.col('End')] = 'Jul 31, 2026'; });
  editBar(wb, 'A-002', (row, c) => { row[c.col('Start')] = '12-Jul-2026'; row[c.col('End')] = '2026-08-06'; });
  const d = P.parse(wb), bar = (id) => d.plan.find((b) => b.id === id);
  assert.deepEqual([bar('A-001').start, bar('A-001').end], ['2026-06-15', '2026-07-31']);
  assert.deepEqual([bar('A-002').start, bar('A-002').end], ['2026-07-12', '2026-08-06']);
});

test('quality: pillar missing from POAP_Config, empty topic, milestone without a date', () => {
  const wb = wbPoap();
  editBar(wb, 'A-001', (row, c) => { row[c.col('Pillar')] = 'Pillar Z – Mystery'; });
  editBar(wb, 'A-002', (row, c) => { row[c.col('Topic')] = ''; });
  edit(wb, 'POAP_Milestones', 'ID', (aoa, c) => { aoa[c.dataRows()[0].i][c.col('Date')] = ''; });
  const m = msgs(P.parse(wb));
  assert.ok(has(m, /pillar 'Pillar Z – Mystery' is not listed in POAP_Config/), m.join('\n'));
  assert.ok(has(m, /A-002 .*Topic is empty/), m.join('\n'));
  assert.ok(has(m, /Milestone 'Release 1 SIT start' has no date and is not drawn/), m.join('\n'));
});

test('parse: percentages accept 0-1, 0-100 and "%" text; status and RAG are canonicalised', () => {
  const wb = wbPoap();
  editBar(wb, 'A-001', (row, c) => { row[c.col('% Complete')] = '50%'; row[c.col('Status')] = 'in progress'; row[c.col('RAG')] = 'amber'; });
  editBar(wb, 'A-002', (row, c) => { row[c.col('% Complete')] = 75; row[c.col('Status')] = 'WIP'; row[c.col('RAG')] = 'R'; });
  editBar(wb, 'A-003', (row, c) => { row[c.col('% Complete')] = 0.25; row[c.col('Status')] = 'late'; row[c.col('RAG')] = 'purple'; });
  const d = P.parse(wb), bar = (id) => d.plan.find((b) => b.id === id);
  assert.deepEqual([bar('A-001').pct, bar('A-001').status, bar('A-001').rag], [0.5, 'In Progress', 'Amber']);
  assert.deepEqual([bar('A-002').pct, bar('A-002').status, bar('A-002').rag], [0.75, 'In Progress', 'Red']);
  assert.deepEqual([bar('A-003').pct, bar('A-003').status, bar('A-003').rag], [0.25, 'Delayed', '']);
});

test('parse: rows marked EXAMPLE in Notes are skipped', () => {
  const wb = wbPoap();
  edit(wb, 'POAP_Plan', 'ID', (aoa, c) => { c.append({ ID: 'EX-1', Pillar: 'Pillar A – Returns', Topic: 'T', 'Activity Type': 'SIT', Item: 'example bar', Start: '2026-09-01', End: '2026-09-02', Notes: 'EXAMPLE – delete' }); });
  assert.equal(P.parse(wb).plan.length, sample.plan.length);
});

test('blank POAP template detects and parses to an empty plan', () => {
  const wb = templateWb('ATS_POAP_Template.xlsx');
  assert.equal(P.detect(wb), true);
  const d = P.parse(wb);
  assert.equal(d.plan.length, 0);
  assert.equal(d.milestones.length, 0);
  assert.deepEqual(P.quality(d), []);
  assert.ok(d.config.pillars.length > 0 && d.config.types.length > 0, 'falls back to default pillars / types');
});

test('detect / parse: non-POAP workbooks are rejected, optional sheets may be absent', () => {
  assert.equal(P.detect(templateWb('ATS_RAID_Log_Template.xlsx')), false);
  assert.equal(P.parse(templateWb('ATS_RAID_Log_Template.xlsx')), null);
  const wb = dropSheet(dropSheet(wbPoap(), 'P2P_Matrix'), 'POAP_Milestones');
  const d = P.parse(wb);
  assert.equal(d.plan.length, sample.plan.length);
  assert.deepEqual([d.milestones.length, d.p2p.length], [0, 0]);
});
