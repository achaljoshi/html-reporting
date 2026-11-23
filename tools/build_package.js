#!/usr/bin/env node
// Build a ready-to-host package: dist/ats-dashboard-<version>.zip (the static site + HOSTING.txt).
// Host it anywhere that serves files (GitLab Pages, IIS, nginx, a shared drive) or just open index.html: no runner/server/build is needed on the receiving side.
// usage: node tools/build_package.js [version]    (default: short git sha or "dev")
const fs = require('fs'), path = require('path'), cp = require('child_process');
const root = path.resolve(__dirname, '..');
let version = process.argv[2];
if (!version) { try { version = cp.execSync('git rev-parse --short HEAD', { cwd: root }).toString().trim(); } catch (e) { version = 'dev'; } }
version = version.replace(/[^\w.-]/g, '_');
const dist = path.join(root, 'dist'), name = 'ats-dashboard-' + version, dir = path.join(dist, name), zip = path.join(dist, name + '.zip');
fs.rmSync(dist, { recursive: true, force: true });
cp.execFileSync('node', [path.join(__dirname, 'build_pages.js'), dir], { stdio: 'inherit' });
fs.writeFileSync(path.join(dir, 'HOSTING.txt'), `ATS Dashboard ${version}

This folder is the finished dashboard. It is plain files: no install, no build, no server-side code.

Quickest: double-click index.html (works offline), then choose your ATS_Program folder.

To host it for a team (the Excel data stays on each person's computer):
  GitLab Pages : copy this folder's contents into a folder called "public" in your GitLab project and commit,
                 with a .gitlab-ci.yml that publishes it (see docs below), or upload it to any web server.
  IIS / nginx / Apache / SharePoint site / shared drive : copy this folder to the web root. Nothing else is needed.

GitLab Pages job that needs no build step and no Node/Docker (any runner that can run a shell):
  pages:
    stage: deploy
    script: [ "test -s public/index.html" ]
    pages: true
    rules: [ { if: '$CI_COMMIT_BRANCH == $CI_DEFAULT_BRANCH' } ]
    # add  tags: [your-runner-tag]  if your runner is tag-restricted

Excel templates and the filling guide are in templates/ and docs/.
`);
cp.execFileSync('zip', ['-qr', zip, name], { cwd: dist, stdio: 'inherit' });
console.log(`packaged ${path.relative(root, zip)} (${(fs.statSync(zip).size / 1024).toFixed(0)} KB)`);
