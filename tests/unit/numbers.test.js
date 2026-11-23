'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { ATS } = require('../helpers/app');
const N = ATS.num;

test('num: numbers, formatted numbers and percentages', () => {
  assert.equal(N.num(5), 5);
  assert.equal(N.num(0), 0);
  assert.equal(N.num('42'), 42);
  assert.equal(N.num(' 7.5 '), 7.5);
  assert.equal(N.num('£1,234.50'), 1234.5);
  assert.equal(N.num('1 000'), 1000);
  assert.equal(N.num('-3'), -3);
  assert.equal(N.num('12%'), 0.12);
  assert.equal(N.num('99.5%'), 0.995);
});

test('num: blanks and text are null, never NaN', () => {
  [null, undefined, '', '   ', 'abc', 'n/a', 'about 4', '%', NaN].forEach((v) => assert.equal(N.num(v), null, JSON.stringify(v)));
});

test('ratio: accepts 0-1 fractions, 0-100 percentages and "%" text', () => {
  assert.equal(N.ratio(0.95), 0.95);
  assert.equal(N.ratio(95), 0.95);
  assert.equal(N.ratio('95%'), 0.95);
  assert.equal(N.ratio('0.5'), 0.5);
  assert.equal(N.ratio(0), 0);
  assert.equal(N.ratio(1), 1, '1 means 100%, not 1%');
  assert.equal(N.ratio('', 0.3), 0.3, 'blank falls back to the default');
  assert.equal(N.ratio('x', 0.3), 0.3);
  assert.equal(N.ratio(null), undefined);
});

test('formatters: pct / gbp / int / plural / str / norm', () => {
  assert.equal(N.pct(0.456), '46%');
  assert.equal(N.pct1(0.4567), '45.7%');
  assert.equal(N.pct(null), '—');
  assert.equal(N.pct(NaN), '—');
  assert.equal(N.gbp(1234567.4), '£1,234,567');
  assert.equal(N.gbp(null), '—');
  assert.equal(N.gbpK(250000), '£250k');
  assert.equal(N.gbpK(1500000), '£1.50m');
  assert.equal(N.int(1234.6), '1,235');
  assert.equal(N.round1(2.26), 2.3);
  assert.equal(N.plural(1, 'defect'), '1 defect');
  assert.equal(N.plural(3, 'defect'), '3 defects');
  assert.equal(N.plural(2, 'risk', 'risks!'), '2 risks!');
  assert.equal(N.str(null), '');
  assert.equal(N.str('  hi '), 'hi');
  assert.equal(N.norm('High Risk – Covered Reqs'), 'highriskcoveredreqs');
});

test('RAG helpers: norm / worst / colour', () => {
  const R = ATS.rag;
  assert.equal(R.norm('green'), 'Green');
  assert.equal(R.norm(' Amber '), 'Amber');
  assert.equal(R.norm('RED'), 'Red');
  assert.equal(R.norm('purple'), '');
  assert.equal(R.norm(null), '');
  assert.equal(R.worst(['Green', 'Amber', 'Green']), 'Amber');
  assert.equal(R.worst(['Green', 'Red', 'Amber']), 'Red');
  assert.equal(R.worst([]), 'Grey');
  assert.equal(R.worst(['Grey', 'Green']), 'Green');
  assert.equal(R.color('nonsense'), R.color('Grey'));
});

test('dom.esc escapes everything that could break out of HTML', () => {
  const esc = ATS.dom.esc;
  assert.equal(esc('<img src=x onerror="alert(1)">'), '&lt;img src=x onerror=&quot;alert(1)&quot;&gt;');
  assert.equal(esc("O'Brien & Co"), 'O&#39;Brien &amp; Co');
  assert.equal(esc(null), '');
  assert.equal(esc(undefined), '');
  assert.equal(esc(0), '0');
});
