'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { ATS } = require('../helpers/app');
const D = ATS.date;

test('toIso: ISO and slash-ordered year-first strings', () => {
  assert.equal(D.toIso('2026-09-25'), '2026-09-25');
  assert.equal(D.toIso('2026/09/25'), '2026-09-25');
  assert.equal(D.toIso('2026-9-5'), '2026-09-05');
  assert.equal(D.toIso('2026-09-25T00:00:00.000Z'), '2026-09-25');
  assert.equal(D.toIso('  2026-09-25  '), '2026-09-25');
});

test('toIso: day-first (UK) numeric dates', () => {
  assert.equal(D.toIso('25/09/2026'), '2026-09-25');
  assert.equal(D.toIso('5/9/2026'), '2026-09-05', 'ambiguous dates are read day-first');
  assert.equal(D.toIso('01/02/2026'), '2026-02-01');
  assert.equal(D.toIso('25.09.2026'), '2026-09-25');
  assert.equal(D.toIso('25-09-2026'), '2026-09-25');
  assert.equal(D.toIso('16/06/26'), '2026-06-16', 'two-digit years are 20xx');
  assert.equal(D.toIso('Tue 16/06/26'), '2026-06-16', 'a leading weekday name is ignored');
});

test('toIso: month-first dates are accepted only when unambiguous', () => {
  assert.equal(D.toIso('9/25/2026'), '2026-09-25');
  assert.equal(D.toIso('12/31/2026'), '2026-12-31');
});

test('toIso: month names', () => {
  assert.equal(D.toIso('Sep 25, 2026'), '2026-09-25');
  assert.equal(D.toIso('September 25, 2026'), '2026-09-25');
  assert.equal(D.toIso('Fri, Sep 25, 2026'), '2026-09-25');
  assert.equal(D.toIso('September 25th, 2026'), '2026-09-25');
  assert.equal(D.toIso('25-Sep-2026'), '2026-09-25');
  assert.equal(D.toIso('25 Sep 2026'), '2026-09-25');
  assert.equal(D.toIso('5 March 26'), '2026-03-05');
});

test('toIso: Excel serial numbers (number and numeric text)', () => {
  const serial = (iso) => Math.round((Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 86400000);
  assert.equal(D.toIso(serial('2026-09-25')), '2026-09-25');
  assert.equal(D.toIso(serial('2024-02-29')), '2024-02-29');
  assert.equal(D.toIso(serial('2026-09-25') + 0.75), '2026-09-25', 'time-of-day fraction is dropped');
  assert.equal(D.toIso(String(serial('2026-09-25'))), '2026-09-25');
  assert.equal(D.toIso(45000), '2023-03-15');
});

test('toIso: Date objects, including timezone skew of up to 12 hours', () => {
  assert.equal(D.toIso(new Date(2026, 8, 25)), '2026-09-25', 'local midnight');
  assert.equal(D.toIso(new Date(2026, 8, 25, 0, 0, 0)), '2026-09-25');
  const nearMidnightBefore = new Date(2026, 8, 24, 23, 59, 50); // the kind of value SheetJS produces for a date-only cell
  assert.equal(D.toIso(nearMidnightBefore), '2026-09-25');
  assert.equal(D.toIso(new Date(NaN)), null);
});

test('toIso: anything that is not a real calendar date is null', () => {
  [null, undefined, '', '   ', 'abc', 'TBC', 'n/a', 'Dates TBC', true, {}, 5, 123, 0, -1, 1e9].forEach((v) => assert.equal(D.toIso(v), null, 'value ' + JSON.stringify(v)));
  assert.equal(D.toIso('2026-25-09'), null, 'month 25 must not be accepted');
  assert.equal(D.toIso('2026-13-01'), null);
  assert.equal(D.toIso('31/02/2026'), null, '31 February');
  assert.equal(D.toIso('32/01/2026'), null);
  assert.equal(D.toIso('29/02/2026'), null, '2026 is not a leap year');
  assert.equal(D.toIso('29/02/2028'), '2028-02-29');
  assert.equal(D.toIso('2026-00-10'), null);
  assert.equal(D.toIso('1/1/1985'), null, 'years before 1990 are rejected');
  assert.equal(D.toIso('1/1/2150'), null, 'years after 2100 are rejected');
});

test('date arithmetic: add / diff / weekEnding / month helpers', () => {
  assert.equal(D.add('2026-12-31', 1), '2027-01-01');
  assert.equal(D.add('2026-03-01', -1), '2026-02-28');
  assert.equal(D.diff('2026-09-01', '2026-09-25'), 24);
  assert.equal(D.diff('2026-09-25', '2026-09-01'), -24);
  assert.equal(D.diff('2026-03-28', '2026-03-30'), 2, 'unaffected by the UK clock change');
  assert.equal(D.dow('2026-09-25'), 5, '25 Sep 2026 is a Friday');
  assert.equal(D.weekEnding('2026-09-21'), '2026-09-25');
  assert.equal(D.weekEnding('2026-09-25'), '2026-09-25');
  assert.equal(D.weekEnding('2026-09-26'), '2026-10-02');
  assert.equal(D.monthEnd('2028-02'), '2028-02-29');
  assert.equal(D.monthEnd('2026-02'), '2026-02-28');
  assert.equal(D.prevMonthKey('2026-01'), '2025-12');
  assert.equal(D.nextMonthKey('2026-12'), '2027-01');
  const q = D.quarterOf('2026-08-15');
  assert.deepEqual([q.start, q.end, q.label], ['2026-07-01', '2026-09-30', 'Q3 2026']);
  assert.equal(D.fmt('2026-09-05'), '5 Sep 2026');
  assert.equal(D.fmt('2026-09-05', false), '5 Sep');
  assert.equal(D.fmt('garbage'), '—');
  assert.equal(D.fmtMonth('2026-09'), 'September 2026');
  assert.equal(D.maxIso('2026-01-01', null), '2026-01-01');
  assert.equal(D.minIso('2026-01-02', '2026-01-01'), '2026-01-01');
});

test('networkdays: Mon-Fri inclusive, like Excel NETWORKDAYS', () => {
  assert.equal(D.networkdays('2026-09-21', '2026-09-25'), 5, 'Mon..Fri');
  assert.equal(D.networkdays('2026-09-21', '2026-09-28'), 6, 'Mon..next Mon');
  assert.equal(D.networkdays('2026-09-25', '2026-09-25'), 1, 'same working day counts as 1');
  assert.equal(D.networkdays('2026-09-26', '2026-09-27'), 0, 'weekend only');
  assert.equal(D.networkdays('2026-09-25', '2026-09-21'), -5, 'reversed range is negative');
  assert.equal(D.networkdays(null, '2026-09-21'), 0);
});

test('networkdays: UK bank holidays from the sample workbook are excluded', () => {
  const { sampleWb } = require('../helpers/app');
  const data = ATS.kpi.parse(sampleWb('Self Assessment', 'ATS_Weekly_Data.xlsx'));
  const hols = new Set(data.holidays);
  ['2026-04-03', '2026-04-06', '2026-08-31', '2026-12-25', '2026-12-28'].forEach((h) => assert.ok(hols.has(h), h + ' should be a bank holiday'));
  // Christmas 2026: Thu 24, Fri 25 (holiday), Mon 28 (Boxing Day substitute), Tue 29
  assert.equal(D.networkdays('2026-12-24', '2026-12-29'), 4);
  assert.equal(D.networkdays('2026-12-24', '2026-12-29', hols), 2);
  // Easter 2026: Good Friday 3 Apr and Easter Monday 6 Apr
  assert.equal(D.networkdays('2026-04-01', '2026-04-08', hols), 4, 'Wed1 Thu2 Tue7 Wed8');
  assert.equal(D.networkdays('2026-08-28', '2026-09-01', hols), 2, 'Fri 28 Aug and Tue 1 Sep (Mon 31 Aug is the summer bank holiday)');
  // a holiday on a weekend must not be subtracted twice
  assert.equal(D.networkdays('2026-09-26', '2026-09-27', new Set(['2026-09-26'])), 0);
});
