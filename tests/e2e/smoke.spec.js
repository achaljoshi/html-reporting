// End-to-end smoke tests: the real app in real Chromium, served by tools/serve.js (see playwright.config.js).
// Helpers in tools/qa_harness.js (QA.sweep, QA.variant, ...) are injected into the page; QA agents can reuse them.
const { test, expect } = require('@playwright/test');
const fs = require('fs');
const path = require('path');

const HARNESS = fs.readFileSync(path.join(__dirname, '..', '..', 'tools', 'qa_harness.js'), 'utf8');
const PORTFOLIOS = ['Child Benefit', 'Customs Declaration Service', 'PAYE', 'Self Assessment', 'VAT'];

// Opens the app and collects every uncaught page error / console.error for the test to assert on.
async function openApp(page) {
  const errors = [];
  page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));
  page.on('console', (m) => { if (m.type() === 'error') errors.push('console.error: ' + m.text()); });
  await page.goto('/index.html');
  await page.waitForSelector('.nav-item');
  await page.waitForFunction(() => window.ATS && ATS.hasData() && document.querySelector('#view-programme').innerText.length > 50);
  return errors;
}
const inject = (page) => page.addScriptTag({ content: HARNESS });
const go = async (page, view, title) => {
  await page.locator('.nav-item[data-view="' + view + '"]').click();
  await expect(page.locator('#view-title')).toHaveText(title);
  await expect(page.locator('#view-' + view)).toBeVisible();
};

test('app loads from the server with the five sample portfolios and no page errors', async ({ page }) => {
  const errors = await openApp(page);
  await expect(page).toHaveTitle(/ATS Program/);
  expect(await page.evaluate(() => ATS.portfolioNames())).toEqual(PORTFOLIOS);
  await expect(page.locator('#program-status')).toContainText('Demo data loaded');
  // the portfolio picker (visible on the single-portfolio sections) lists all five
  await go(page, 'monthly', 'Monthly Council');
  await expect(page.locator('#portfolio-select option')).toHaveText(PORTFOLIOS);
  expect(errors).toEqual([]);
});

test('QA sweep of every Monthly and Weekly page for one portfolio finds nothing', async ({ page }) => {
  test.setTimeout(150 * 1000);
  const errors = await openApp(page);
  await inject(page);
  const findings = await page.evaluate(() => QA.sweep('Self Assessment'));
  const checked = await page.evaluate(() => QA.checked);
  expect(checked, 'the sweep must actually have inspected pages').toBeGreaterThan(20);
  expect(findings).toEqual([]);
  expect(await page.evaluate(() => QA.errors)).toEqual([]);
  expect(errors).toEqual([]);
});

test('section navigation works with real clicks and keeps the portfolio picker in sync', async ({ page }) => {
  const errors = await openApp(page);
  await expect(page.locator('#view-title')).toHaveText('Programme Overview');
  await expect(page.locator('#portfolio-wrap')).toBeHidden();

  await go(page, 'monthly', 'Monthly Council');
  await expect(page.locator('#portfolio-wrap')).toBeVisible();
  await expect(page.locator('#view-monthly .sx-head h2')).toContainText(/\w+ 20\d\d/);
  await expect(page.locator('#view-monthly .sx-rail button, #view-monthly .sx-rail a').first()).toBeVisible();

  await go(page, 'weekly', 'Weekly Report');
  await expect(page.locator('#view-weekly .sx-head h2')).toContainText('Week ending');

  await page.selectOption('#portfolio-select', 'Self Assessment');
  await go(page, 'poap', 'POAP — Plan on a Page');
  await expect(page.locator('#poap-root')).toBeVisible();

  await go(page, 'programme', 'Programme Overview');
  await expect(page.locator('#view-programme')).toContainText('Self Assessment');
  expect(errors).toEqual([]);
});

test('the Monthly Council period can be stepped with the on-page controls', async ({ page }) => {
  const errors = await openApp(page);
  await go(page, 'monthly', 'Monthly Council');
  const head = page.locator('#view-monthly .sx-head h2');
  const before = await head.innerText();
  const prev = page.locator('#view-monthly .sx-controls button[title="Previous"]');
  await expect(prev).toBeEnabled();
  await prev.click();
  await expect(head).not.toHaveText(before);
  await page.locator('#view-monthly .sx-controls button[title="Next"]').click();
  await expect(head).toHaveText(before);
  expect(errors).toEqual([]);
});

test('Export PDF (Monthly Council) produces a PDF download', async ({ page }) => {
  test.setTimeout(150 * 1000);
  const errors = await openApp(page);
  await go(page, 'monthly', 'Monthly Council');
  const [download] = await Promise.all([page.waitForEvent('download', { timeout: 120 * 1000 }), page.locator('#btn-pdf').click()]);
  expect(download.suggestedFilename()).toMatch(/^Monthly_Council_.*\.pdf$/);
  const file = await download.path();
  const head = fs.readFileSync(file).subarray(0, 5).toString('latin1');
  expect(head).toBe('%PDF-');
  expect(fs.statSync(file).size).toBeGreaterThan(20 * 1024);
  expect(errors).toEqual([]);
});

// minimal zip reader: names of all entries from the central directory (no dependency needed)
function zipEntries(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 66000); i--) { if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; } }
  if (eocd < 0) throw new Error('not a zip file (no end-of-central-directory record)');
  const count = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const names = [];
  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error('corrupt central directory');
    const len = buf.readUInt16LE(p + 28), extra = buf.readUInt16LE(p + 30), comment = buf.readUInt16LE(p + 32);
    names.push(buf.toString('utf8', p + 46, p + 46 + len));
    p += 46 + len + extra + comment;
  }
  return names;
}

test('Export PowerPoint produces a valid .pptx (zip with slides)', async ({ page }) => {
  test.setTimeout(150 * 1000);
  const errors = await openApp(page);
  await go(page, 'monthly', 'Monthly Council');
  const [download] = await Promise.all([page.waitForEvent('download', { timeout: 120 * 1000 }), page.getByRole('button', { name: 'Export PowerPoint' }).click()]);
  expect(download.suggestedFilename()).toMatch(/^Monthly_Council_.*\.pptx$/);
  const buf = fs.readFileSync(await download.path());
  expect(buf.subarray(0, 4).toString('latin1')).toBe('PK\x03\x04');
  const names = zipEntries(buf);
  expect(names).toContain('[Content_Types].xml');
  expect(names).toContain('ppt/presentation.xml');
  expect(names).toContain('ppt/slides/slide1.xml');
  expect(names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length).toBeGreaterThan(3);
  expect(errors).toEqual([]);
});

test('POAP roadmap shows activity rows and bars for a portfolio that has a plan', async ({ page }) => {
  const errors = await openApp(page);
  await go(page, 'monthly', 'Monthly Council');
  await page.selectOption('#portfolio-select', 'Self Assessment');
  await go(page, 'poap', 'POAP — Plan on a Page');
  const rows = page.locator('#poap-root .poap-rl');
  await expect(rows.first()).toBeVisible();
  expect(await rows.count()).toBeGreaterThan(5);
  expect(await page.locator('#poap-root .poap-bar').count()).toBeGreaterThan(10);
  expect(errors).toEqual([]);
});

test('a portfolio without a POAP says so instead of breaking', async ({ page }) => {
  const errors = await openApp(page);
  await go(page, 'monthly', 'Monthly Council');
  await page.selectOption('#portfolio-select', 'VAT');
  await go(page, 'poap', 'POAP — Plan on a Page');
  await expect(page.locator('#view-poap')).toContainText('No POAP for VAT');
  expect(errors).toEqual([]);
});
