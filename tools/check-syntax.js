#!/usr/bin/env node
// `npm run build` — there is no bundler, so "build" = syntax-check every shipped script with `node --check`.
// Covers assets/js/*.js and tools/*.js. Exits non-zero if any file fails.
const fs = require('fs'), path = require('path'), { spawnSync } = require('child_process');
const root = path.resolve(__dirname, '..');
const dirs = ['assets/js', 'tools'];
const files = [];
dirs.forEach((d) => {
  fs.readdirSync(path.join(root, d)).sort().forEach((f) => { if (/\.js$/.test(f)) files.push(path.join(d, f)); });
});
let failed = 0;
files.forEach((f) => {
  const r = spawnSync(process.execPath, ['--check', path.join(root, f)], { encoding: 'utf8' });
  if (r.status === 0) console.log('ok   ' + f);
  else { failed++; console.error('FAIL ' + f + '\n' + (r.stderr || r.stdout || '').trim()); }
});
console.log('\n' + (files.length - failed) + '/' + files.length + ' files passed syntax check');
process.exit(failed ? 1 : 0);
