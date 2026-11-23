const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs'), os = require('os'), path = require('path'), cp = require('child_process');

test('GitLab Pages build contains the app and none of the private or test folders', () => {
  const out = fs.mkdtempSync(path.join(os.tmpdir(), 'pages-'));
  cp.execFileSync('node', [path.join(__dirname, '..', '..', 'tools', 'build_pages.js'), out], { stdio: 'pipe' });
  for (const must of ['index.html', 'assets/js/core.js', 'assets/lib/xlsx.full.min.js', 'assets/js/sample-data.js', 'templates/ATS_Weekly_Data_Template.xlsx']) assert.ok(fs.existsSync(path.join(out, must)), must);
  for (const never of ['ATS_Program', 'local_private', 'node_modules', 'tests', '.git', 'test-results', 'samples']) assert.ok(!fs.existsSync(path.join(out, never)), never + ' must not be published');
  fs.rmSync(out, { recursive: true, force: true });
});
