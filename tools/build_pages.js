#!/usr/bin/env node
// Assemble the static site for GitLab Pages into ./public (only what the app needs at runtime; no tests, no private folders).
// usage: node tools/build_pages.js [outDir]
const fs = require('fs'), path = require('path');
const root = path.resolve(__dirname, '..'), out = path.resolve(root, process.argv[2] || 'public');
const INCLUDE = ['index.html', 'assets', 'templates', 'docs', 'README.md'];   // templates + docs = downloadable Excel templates and the filling guide
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });
for (const item of INCLUDE) fs.cpSync(path.join(root, item), path.join(out, item), { recursive: true });
fs.writeFileSync(path.join(out, '.nojekyll'), '');
// every local file index.html loads must be in the output, or the site would be half-broken
const html = fs.readFileSync(path.join(out, 'index.html'), 'utf8');
const refs = [...html.matchAll(/(?:src|href)="(?!https?:|#|data:)([^"]+)"/g)].map((m) => m[1].split('?')[0]);
const missing = refs.filter((r) => !fs.existsSync(path.join(out, r)));
if (missing.length) { console.error('missing from output:', missing); process.exit(1); }
const count = (d) => fs.readdirSync(d, { withFileTypes: true }).reduce((n, e) => n + (e.isDirectory() ? count(path.join(d, e.name)) : 1), 0);
console.log(`built ${out}: ${count(out)} files, ${refs.length} index.html references all present`);
