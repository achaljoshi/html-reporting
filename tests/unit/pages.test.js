'use strict';
// public/ (what GitLab Pages serves) must match the app sources: regenerate with `npm run pages` after changing the app.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const root = path.resolve(__dirname, '..', '..');
const INCLUDE = ['index.html', 'assets', 'templates', 'docs', 'README.md'];
const list = (base, rel) => {
  const abs = path.join(base, rel);
  if (!fs.existsSync(abs)) return [];
  return fs.statSync(abs).isDirectory() ? fs.readdirSync(abs).flatMap((n) => list(base, path.join(rel, n))) : [rel];
};
test('public/ is in sync with the app sources (run `npm run pages` if this fails)', () => {
  const src = INCLUDE.flatMap((i) => list(root, i)).sort();
  const pub = list(path.join(root, 'public'), '.').map((f) => f.replace(/^\.[\\/]/, '')).filter((f) => f !== '.nojekyll').sort();
  assert.deepStrictEqual(pub, src, 'file list differs');
  const stale = src.filter((f) => !fs.readFileSync(path.join(root, f)).equals(fs.readFileSync(path.join(root, 'public', f))));
  assert.deepStrictEqual(stale, [], 'stale files in public/');
});
