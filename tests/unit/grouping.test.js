'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { ATS, XLSX, SAMPLE_DIR, kpiSamples, readWb } = require('../helpers/app');

// --- tiny synthetic "parsed workbook" items (the shape built by ATS.loadFiles / tools/build_sample_js.js)
const kpi = (portfolio, dir, o) => ({ kind: 'kpi', parsed: { config: { portfolio: portfolio || '' }, snapshots: [], defects: [] }, dir: dir || '', fileName: (o && o.fileName) || 'ATS_Weekly_Data.xlsx', ts: (o && o.ts) || 1, sz: o && o.sz != null ? o.sz : 10 });
const poap = (program, dir, o) => ({ kind: 'poap', parsed: { config: { program: program || '' }, plan: [] }, dir: dir || '', fileName: (o && o.fileName) || 'ATS_POAP.xlsx', ts: (o && o.ts) || 1, sz: o && o.sz != null ? o.sz : 5 });
const raid = (project, dir, o) => ({ kind: 'raid', parsed: { info: { project: project || '' }, risks: [], issues: [] }, dir: dir || '', fileName: (o && o.fileName) || 'ATS_RAID_Log.xlsx', ts: (o && o.ts) || 1, sz: o && o.sz != null ? o.sz : 3 });
const names = (groups) => groups.map((g) => g.name);

test('portfolioName precedence: Config name > folder > file name > template > program name', () => {
  const P = ATS.portfolioName;
  assert.equal(P({ config: { portfolio: 'PAYE' } }, 'x.xlsx', 'Some Folder'), 'PAYE');
  assert.equal(P({ config: { portfolio: '' } }, 'ATS_Weekly_Data.xlsx', 'Root/VAT'), 'VAT', 'last folder segment');
  assert.equal(P({ config: {} }, 'ATS_Weekly_Data - Child Benefit.xlsx', ''), 'Child Benefit');
  assert.equal(P({ config: {} }, 'ATS_Weekly_Data_Template.xlsx', ''), '(blank template)');
  assert.equal(P({ config: { programName: 'HMRC ATS – Customs' } }, 'weekly.xlsx', ''), 'Customs');
  assert.equal(P({ config: { portfolio: '<Portfolio Name>' } }, 'weekly.xlsx', 'Fallback'), 'Fallback', 'unfilled <placeholder> is not a name');
});

test('groupPortfolios: one portfolio per folder, sorted by folder', () => {
  const g = ATS.groupPortfolios([kpi('B', 'B'), kpi('A', 'A')]);
  assert.deepEqual(names(g), ['A', 'B']);
  assert.ok(!g[0].poap && !g[0].raid, 'POAP and RAID are optional');
});

test('groupPortfolios: POAP and RAID in the same folder attach to that portfolio only', () => {
  const g = ATS.groupPortfolios([kpi('A', 'A'), poap('A', 'A'), raid('A', 'A'), kpi('B', 'B')]);
  const a = g.find((x) => x.name === 'A'), b = g.find((x) => x.name === 'B');
  assert.ok(a.poap && a.raid);
  assert.ok(!b.poap && !b.raid);
  assert.equal(a._items.length, 3);
});

test('groupPortfolios: blank templates are ignored whenever a real workbook of that kind exists', () => {
  const g = ATS.groupPortfolios([kpi('Real', 'Real'), kpi('', '', { fileName: 'ATS_Weekly_Data_Template.xlsx', sz: 0 })]);
  assert.deepEqual(names(g), ['Real']);
  // ...but a lone blank template is still shown (so a brand-new user sees something)
  const only = ATS.groupPortfolios([kpi('', '', { fileName: 'ATS_Weekly_Data_Template.xlsx', sz: 0 })]);
  assert.deepEqual(names(only), ['(blank template)']);
});

test('groupPortfolios: blank POAP/RAID templates never replace a real one', () => {
  const g = ATS.groupPortfolios([kpi('A', 'A'), poap('A', 'A'), poap('A', 'A', { fileName: 'ATS_POAP_Template.xlsx', sz: 0, ts: 99 }), raid('A', 'A', { sz: 0, ts: 99 })]);
  assert.equal(g.length, 1);
  assert.equal(g[0].poap.plan.length, 0);
  assert.equal(g[0]._items.filter((i) => i.kind === 'poap').length, 1);
  assert.equal(g[0]._items.find((i) => i.kind === 'poap').fileName, 'ATS_POAP.xlsx');
});

test('groupPortfolios: two workbooks for one portfolio — the most recently saved wins', () => {
  const older = kpi('A', 'A', { fileName: 'old.xlsx', ts: 100 }), newer = kpi('A', 'A', { fileName: 'new.xlsx', ts: 200 });
  const g = ATS.groupPortfolios([older, newer]);
  assert.equal(g.length, 1);
  assert.equal(g[0]._items[0].fileName, 'new.xlsx');
  // input order does not matter
  assert.equal(ATS.groupPortfolios([newer, older])[0]._items[0].fileName, 'new.xlsx');
});

test('groupPortfolios: a workbook with data beats a newer empty one', () => {
  const withData = kpi('A', 'A', { fileName: 'data.xlsx', ts: 100, sz: 12 }), empty = kpi('A', 'A', { fileName: 'empty.xlsx', ts: 500, sz: 0 });
  // empty one is filtered out as a template because real data exists
  assert.equal(ATS.groupPortfolios([withData, empty])[0]._items[0].fileName, 'data.xlsx');
});

test('groupPortfolios: same portfolio name in different folders gets a numeric suffix', () => {
  const g = ATS.groupPortfolios([kpi('Dup', 'one'), kpi('Dup', 'two')]);
  assert.deepEqual(names(g).sort(), ['Dup', 'Dup (2)']);
});

test('groupPortfolios: a flat folder with several weekly workbooks matches POAP/RAID by name', () => {
  const g = ATS.groupPortfolios([
    kpi('Alpha', '', { fileName: 'ATS_Weekly_Data - Alpha.xlsx' }), kpi('Beta', '', { fileName: 'ATS_Weekly_Data - Beta.xlsx' }),
    poap('Beta programme', '', { fileName: 'ATS_POAP - Beta.xlsx' }), raid('Alpha', '', { fileName: 'raid.xlsx' }),
  ]);
  const alpha = g.find((x) => x.name === 'Alpha'), beta = g.find((x) => x.name === 'Beta');
  assert.ok(beta.poap && !beta.raid, 'POAP matched by file name');
  assert.ok(alpha.raid && !alpha.poap, 'RAID matched by the project name inside the file');
});

test('groupPortfolios: an unmatched POAP in a flat folder with several portfolios is not guessed', () => {
  const g = ATS.groupPortfolios([kpi('Alpha', ''), kpi('Beta', '', { fileName: 'b.xlsx' }), poap('Gamma', '', { fileName: 'plan.xlsx' })]);
  assert.ok(g.every((p) => !p.poap));
});

test('groupPortfolios: POAP/RAID in a folder without a weekly workbook attach by name, or to the only portfolio lacking one', () => {
  const byName = ATS.groupPortfolios([kpi('Alpha', 'Alpha'), kpi('Beta', 'Beta'), poap('Beta', 'Plans', { fileName: 'ATS_POAP - Beta.xlsx' })]);
  assert.ok(byName.find((p) => p.name === 'Beta').poap);
  assert.ok(!byName.find((p) => p.name === 'Alpha').poap);
  const lone = ATS.groupPortfolios([kpi('Alpha', 'Alpha', { }), poap('Whatever', 'Plans', { fileName: 'plan.xlsx' })]);
  assert.ok(lone[0].poap, 'the single portfolio without a POAP receives the orphan');
  const ambiguous = ATS.groupPortfolios([kpi('Alpha', 'Alpha'), kpi('Beta', 'Beta'), raid('Whatever', 'Logs', { fileName: 'raid.xlsx' })]);
  assert.ok(ambiguous.every((p) => !p.raid), 'two candidates: do not guess');
});

test('groupPortfolios: no weekly workbook means no portfolios', () => {
  assert.deepEqual(ATS.groupPortfolios([poap('A', 'A'), raid('A', 'A')]), []);
  assert.deepEqual(ATS.groupPortfolios([]), []);
});

test('real sample folder groups into five portfolios; only Self Assessment has POAP + RAID', () => {
  const items = [];
  (function walk(dir, rel) {
    fs.readdirSync(dir, { withFileTypes: true }).forEach((e) => {
      if (e.isDirectory()) return walk(path.join(dir, e.name), rel ? rel + '/' + e.name : e.name);
      if (!/\.xlsx$/i.test(e.name)) return;
      const wb = readWb(path.join(dir, e.name));
      const kind = ATS.kpi.detect(wb) ? 'kpi' : ATS.poap.detect(wb) ? 'poap' : ATS.raid.detect(wb) ? 'raid' : null;
      assert.ok(kind, e.name + ' should be recognised as an ATS workbook');
      const parsed = ATS[kind].parse(wb);
      const sz = kind === 'kpi' ? parsed.snapshots.length + parsed.defects.length : kind === 'poap' ? parsed.plan.length : parsed.risks.length + parsed.issues.length;
      items.push({ kind, parsed, dir: rel, fileName: e.name, ts: 1, sz });
    });
  })(SAMPLE_DIR, '');
  const groups = ATS.groupPortfolios(items);
  assert.deepEqual(names(groups).sort(), ['Child Benefit', 'Customs Declaration Service', 'PAYE', 'Self Assessment', 'VAT']);
  groups.forEach((g) => assert.ok(g.kpi && g.kpi.snapshots.length > 0, g.name + ' has weekly data'));
  const withExtras = groups.filter((g) => g.poap || g.raid).map((g) => g.name);
  assert.deepEqual(withExtras, ['Self Assessment']);
  const sa = groups.find((g) => g.name === 'Self Assessment');
  assert.ok(sa.poap.plan.length > 0 && sa.raid.risks.length > 0);
  assert.equal(kpiSamples().length, 5);
});

test('blank templates are recognised as ATS workbooks of the right kind', () => {
  const { templateWb } = require('../helpers/app');
  assert.ok(ATS.kpi.detect(templateWb('ATS_Weekly_Data_Template.xlsx')));
  assert.ok(ATS.poap.detect(templateWb('ATS_POAP_Template.xlsx')));
  assert.ok(ATS.raid.detect(templateWb('ATS_RAID_Log_Template.xlsx')));
  const random = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(random, XLSX.utils.aoa_to_sheet([['a']]), 'Sheet1');
  assert.ok(!ATS.kpi.detect(random) && !ATS.poap.detect(random) && !ATS.raid.detect(random));
});
