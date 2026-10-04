'use strict';
// Loads the browser modules into Node the same way tools/build_sample_js.js does, and offers workbook helpers
// for building edge-case data in code (no binary fixtures are committed: we load a sample workbook and mutate it).
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
global.window = global;
global.XLSX = require(path.join(ROOT, 'assets/lib/xlsx.full.min.js'));
['core.js', 'kpi-data.js', 'raid.js', 'poap.js'].forEach((f) => require(path.join(ROOT, 'assets/js', f)));
const ATS = global.ATS;
const XLSX = global.XLSX;

const SAMPLE_DIR = path.join(ROOT, 'samples', 'ATS_Program_SAMPLE');
const TEMPLATE_DIR = path.join(ROOT, 'templates');

function readWb(file) {
  return XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: true });
}

// every samples/ATS_Program_SAMPLE/**/<name> found, as { portfolio (folder name), file }
function sampleFiles(name) {
  const out = [];
  fs.readdirSync(SAMPLE_DIR, { withFileTypes: true }).forEach((e) => {
    if (e.isDirectory() && fs.existsSync(path.join(SAMPLE_DIR, e.name, name))) out.push({ folder: e.name, file: path.join(SAMPLE_DIR, e.name, name) });
  });
  return out;
}
const kpiSamples = () => sampleFiles('ATS_Weekly_Data.xlsx');
const sampleWb = (folder, name) => readWb(path.join(SAMPLE_DIR, folder, name));
const templateWb = (name) => readWb(path.join(TEMPLATE_DIR, name));

// ---- workbook mutation helpers -------------------------------------------------------------
// Edit a sheet as an array of arrays. fn(aoa, ctx) where ctx = { hdr (header row index), col(name) -> column index }.
// The header row is the first row whose first cell equals `firstHeader` (case-insensitive).
function edit(wb, sheet, firstHeader, fn) {
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, raw: true, defval: '' });
  const hdr = firstHeader == null ? -1 : aoa.findIndex((r) => String(r[0]).trim().toLowerCase() === String(firstHeader).toLowerCase());
  if (firstHeader != null && hdr < 0) throw new Error('header "' + firstHeader + '" not found on ' + sheet);
  const col = (name) => {
    const i = aoa[hdr].findIndex((c) => String(c).trim().toLowerCase().startsWith(String(name).toLowerCase()));
    if (i < 0) throw new Error('column "' + name + '" not found on ' + sheet);
    return i;
  };
  const ctx = {
    hdr, col,
    dataRows: () => aoa.slice(hdr + 1).map((r, i) => ({ r, i: hdr + 1 + i })).filter((x) => x.r.some((c) => c !== '' && c != null)),
    set: (rowIdx, name, value) => { aoa[rowIdx][col(name)] = value; },
    // append a row (object keyed by column-name prefix) into the first blank row under the table
    append: (obj) => {
      let i = hdr + 1;
      while (i < aoa.length && aoa[i].some((c) => c !== '' && c != null)) i++;
      if (i >= aoa.length) aoa.push(new Array(aoa[hdr].length).fill(''));
      Object.keys(obj).forEach((k) => { aoa[i][col(k)] = obj[k]; });
      return i;
    },
  };
  fn(aoa, ctx);
  wb.Sheets[sheet] = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
  return wb;
}

function dropSheet(wb, name) {
  wb.SheetNames = wb.SheetNames.filter((n) => n !== name);
  delete wb.Sheets[name];
  return wb;
}

// most recent Friday on or before the real clock (so "fresh" test data never goes stale)
function lastFriday(offsetWeeks) {
  const D = ATS.date;
  let d = D.todayIso();
  while (D.dow(d) !== 5) d = D.add(d, -1);
  return D.add(d, -7 * (offsetWeeks || 0));
}

module.exports = { ROOT, SAMPLE_DIR, TEMPLATE_DIR, ATS, XLSX, readWb, sampleFiles, kpiSamples, sampleWb, templateWb, edit, dropSheet, lastFriday };
