// @ts-check
// Playwright config for the offline dashboard. The app itself needs no server (index.html works from file://);
// the tests serve the repo root with tools/serve.js so downloads / fetches behave like a normal website.
const { defineConfig, devices } = require('@playwright/test');

const PORT = process.env.PORT || 8765;
const BASE_URL = 'http://localhost:' + PORT;

module.exports = defineConfig({
  testDir: 'tests',
  // only the e2e smoke tests and anything the QA agent generates; tests/unit is run by `node --test`
  testMatch: ['e2e/**/*.spec.js', 'ai-generated/**/*.spec.js'],
  fullyParallel: false,
  workers: 1,
  retries: 0,
  timeout: 90 * 1000,
  expect: { timeout: 10 * 1000 },
  reporter: [
    ['list'],
    ['html', { outputFolder: 'playwright-report', open: 'never' }],
    ['junit', { outputFile: 'junit-qa.xml' }],
  ],
  use: {
    baseURL: BASE_URL,
    viewport: { width: 1440, height: 900 },
    acceptDownloads: true,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  webServer: {
    command: 'node tools/serve.js',
    url: BASE_URL,
    reuseExistingServer: true,
    timeout: 30 * 1000,
    env: { PORT: String(PORT) },
  },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'], viewport: { width: 1440, height: 900 } } }],
});
