// Pictures and Outlook emails embedded in the demo workbooks: chips next to the notes, the Evidence viewer, the library page.
const { test, expect } = require('@playwright/test');

async function open(page) {
  const errors = [];
  page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));
  // the sandboxed email frame blocks scripts on purpose; Playwright's own injected helper script triggers this message, the page does not
  page.on('console', (m) => { if (m.type() === 'error' && !/Blocked script execution in 'about:srcdoc'/.test(m.text())) errors.push('console.error: ' + m.text()); });
  await page.goto('/index.html');
  await page.waitForFunction(() => window.ATS && ATS.hasData());
  await page.evaluate(() => ATS.setPortfolio('Self Assessment'));
  return errors;
}
const monthlyPage = async (page, id) => {
  await page.locator('.nav-item[data-view="monthly"]').click();
  await page.locator('#m-rail button[data-page="' + id + '"]').click();
};

test('a risk with a picture and an email shows a chip; the viewer opens the picture, then the email', async ({ page }) => {
  const errors = await open(page);
  await monthlyPage(page, 'risks');
  // R-001 carries a picture and an email (R-003 carries one email): pick the chip by what is inside it
  const chip = page.locator('#m-stage .att-chip[title*="environment-health.png"]').first();
  await expect(chip).toBeVisible();
  await expect(chip).toContainText('2');
  await chip.click();
  const viewer = page.locator('#att-viewer.open');
  await expect(viewer).toBeVisible();
  await expect(viewer.locator('.att-name')).toHaveText('environment-health.png');
  await expect(viewer.locator('.att-stage-img img')).toBeVisible();
  await expect(viewer.locator('.att-eyebrow')).toHaveText('EVIDENCE');
  // the second file is the email: a readable card with sender, recipients and the body rendered in a locked-down frame
  await viewer.locator('.att-item').nth(1).click();
  await expect(viewer.locator('.att-mail-hd h3')).toContainText('Integration environment outage');
  await expect(viewer.locator('.att-mail-hd')).toContainText('Environment Team');
  const frame = viewer.locator('iframe.att-mail-frame');
  await expect(frame).toHaveAttribute('sandbox', /allow-popups/);
  expect(await frame.getAttribute('sandbox')).not.toMatch(/allow-scripts|allow-same-origin/);
  await expect(page.frameLocator('iframe.att-mail-frame').locator('body')).toContainText('unavailable from 09:10 to 16:40');
  // keyboard: Escape closes
  await page.keyboard.press('Escape');
  await expect(page.locator('#att-viewer.open')).toHaveCount(0);
  expect(errors).toEqual([]);
});

test('downloads: the original .msg and a .eml export of the email', async ({ page }) => {
  await open(page);
  await monthlyPage(page, 'risks');
  await page.locator('#m-stage .att-chip[title*="environment-health.png"]').first().click();
  const viewer = page.locator('#att-viewer.open');
  await viewer.locator('.att-item').nth(1).click();
  const [msg] = await Promise.all([page.waitForEvent('download'), viewer.locator('.att-top [data-dl="orig"]').click()]);
  expect(msg.suggestedFilename()).toMatch(/\.msg$/);
  const [eml] = await Promise.all([page.waitForEvent('download'), viewer.locator('[data-dl="eml"]').click()]);
  expect(eml.suggestedFilename()).toMatch(/\.eml$/);
  const text = require('fs').readFileSync(await eml.path(), 'utf8');
  expect(text).toContain('Subject:');
  expect(text).toContain('environment-health.png');
});

test('an email with a screenshot opens the screenshot in a nested view and Esc steps back', async ({ page }) => {
  await open(page);
  await monthlyPage(page, 'evidence');
  await page.locator('#m-ev-q').fill('SA-106');
  const card = page.locator('#m-stage .att-card', { hasText: /submission timeout/i }).first();
  await card.click();
  const viewer = page.locator('#att-viewer.open');
  await expect(viewer.locator('.att-mail-hd h3')).toContainText('SA-106');
  await viewer.locator('.att-file button[data-nest]').first().click();
  await expect(viewer.locator('.att-backbtn')).toBeVisible();
  await expect(viewer.locator('.att-stage-img img')).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(viewer.locator('.att-mail-hd h3')).toBeVisible();      // back on the email
  await page.keyboard.press('Escape');
  await expect(page.locator('#att-viewer.open')).toHaveCount(0);
});

test('Evidence library lists everything for the portfolio, filters, and opens the gallery', async ({ page }) => {
  const errors = await open(page);
  await monthlyPage(page, 'evidence');
  await expect(page.locator('#m-stage .att-card').first()).toBeVisible();
  const total = await page.locator('#m-stage .att-card').count();
  expect(total).toBeGreaterThanOrEqual(10);
  await page.locator('#m-ev-kind button[data-k="msg"]').click();
  const emails = await page.locator('#m-stage .att-card').count();
  expect(emails).toBeGreaterThan(0);
  expect(emails).toBeLessThan(total);
  await page.locator('#m-ev-q').fill('Vendor build 3.2');
  await expect(page.locator('#m-stage .att-card')).toHaveCount(1);
  await page.locator('#m-stage .att-card').first().click();
  await expect(page.locator('#att-viewer.open .att-mail-hd h3')).toContainText('Vendor build');
  await page.keyboard.press('Escape');
  expect(errors).toEqual([]);
});

test('a portfolio without attachments says how to add them instead of showing an empty page', async ({ page }) => {
  await open(page);
  await page.evaluate(() => ATS.setPortfolio('PAYE'));
  await monthlyPage(page, 'evidence');
  await expect(page.locator('#m-stage')).toContainText('No attachments found');
  await expect(page.locator('#m-stage')).toContainText('Insert');
});

test('POAP: an activity with an attachment lists it in the detail drawer', async ({ page }) => {
  await open(page);
  await page.locator('.nav-item[data-view="poap"]').click();
  await page.waitForSelector('#poap-root .poap-rl-item');
  await page.evaluate(() => ATS.poap.openDetails('bar', 'A-018'));
  const row = page.locator('.poap-drawer .att-row').first();
  await expect(row).toBeVisible();
  await row.click();
  await expect(page.locator('#att-viewer.open .att-name')).toHaveText('sit-burn-up.png');
});

test('attachments survive a page reload (saved with the data)', async ({ page }) => {
  await open(page);
  await page.evaluate(async () => { const b = await ATS.att.get(Object.keys(ATS.att.registry)[0]); return !!b; });
  const before = await page.evaluate(() => Object.keys(ATS.att.registry).length);
  expect(before).toBeGreaterThanOrEqual(10);
});

test('security: markup inside an email subject, sender, file name and body cannot run in the dashboard', async ({ page }) => {
  await open(page);
  await page.evaluate(() => {
    window.__pwn = 0;
    const evil = '<img src=x onerror="window.__pwn=1">';
    const em = { subject: evil, from: { name: evil, email: 'a@b.c' }, to: [{ name: evil, email: '' }], cc: [], bcc: [], date: '2026-01-01T00:00:00Z', bodyText: evil, bodyHtml: '<script>parent.__pwn=2<\/script>' + evil, headers: evil,
      attachments: [{ name: evil + '.png', mime: 'image/png', size: 70, b64: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==' }] };
    ATS.att.registry.evil = { id: 'evil', name: evil + '.msg', kind: 'msg', mime: 'application/vnd.ms-outlook', size: 5, b64: 'AAAA', email: em, owner: evil, sheet: evil, row: 1, portfolio: 'Self Assessment' };
    ATS.att.open(['evil'], { label: evil });
  });
  await expect(page.locator('#att-viewer.open .att-mail-hd h3')).toBeVisible();
  await page.locator('#att-viewer [data-tab="text"]').click();
  await page.locator('#att-viewer [data-tab="headers"]').click();
  await page.locator('#att-viewer [data-tab="message"]').click();
  await page.waitForTimeout(400);
  expect(await page.evaluate(() => window.__pwn)).toBe(0);                    // no handler ran, in the page or from the sandboxed frame
  expect(await page.locator('#att-viewer img[onerror], #att-viewer script').count()).toBe(0);
  await expect(page.locator('#att-viewer .att-mail-hd h3')).toContainText('<img src=x');   // shown as text
});
