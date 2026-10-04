# AGENTS.md — contract for AI agents working in this repo

This repo is the **ATS Executive Dashboard**: an offline, no-server HTML reporting app. A Development agent builds/fixes
features and a QA agent tests them; both are driven by the `azure-ai-automation` orchestrator, which reads
[`.agent/project.json`](.agent/project.json). Humans own the guard-rails listed under *Protected paths*.

## Architecture (read this first)
1. Plain browser scripts — **no bundler, no framework, no build step**. `index.html` loads scripts in a fixed order; there are no ES modules.
2. Everything hangs off one global namespace, `window.ATS` (`ATS.date`, `ATS.num`, `ATS.dom`, `ATS.rag`, `ATS.xl`, `ATS.kpi`, `ATS.raid`, `ATS.poap`, `ATS.store`, `ATS.shell` ...). Each file is an IIFE that extends `window.ATS`.
3. Libraries are **bundled** in `assets/lib/` (SheetJS, Chart.js, html2canvas, jsPDF, PptxGenJS). No CDN, no network calls at runtime.
4. **`file://` must keep working**: double-clicking `index.html` is the supported way to run it. Do not add `fetch`/XHR of local files, ES modules, or anything that needs a server. (`npm run serve` exists for tests only.)
5. Data flow: Excel workbooks -> `ATS.loadFiles` -> `ATS.kpi.parse` / `ATS.poap.parse` / `ATS.raid.parse` -> `ATS.groupPortfolios` -> `ATS.store` (persisted in IndexedDB) -> pages.
6. `assets/js/core.js` helpers + store + grouping · `kpi-data.js` KPI parse + all KPI maths (`ATS.kpi.computeAll`) · `raid.js` RAID parse/summary · `poap.js` POAP parse + the whole POAP UI.
7. Pages: `council.js` (section framework), `monthly.js` / `weekly.js` (page lists), `kpi-pages.js` (one visual per KPI), `raid-pages.js`, `programme.js` (all-portfolio overview), `shell.js` (sidebar, routing, loading).
8. Exports: `pptx-export.js` (+ `pptx-template.js`, the client theme — protected), `pdf-export.js`.
9. Demo data is **embedded**: `assets/js/sample-data.js` is generated from `samples/ATS_Program_SAMPLE/**` by `node tools/build_sample_js.js`. Re-run it (and commit the result) whenever a parser changes the parsed shape.
10. **Bump `ATS.VER`** (top of the data store in `core.js`) whenever the *shape* of parsed data changes; it invalidates the browser's cached dataset.
11. Dates are ISO `YYYY-MM-DD` strings everywhere; working-day maths uses `ATS.date.networkdays` with UK bank holidays from the workbook.
12. Column names in workbooks are matched by normalised header text (`ATS.num.norm`), so parsers tolerate reordered/renamed-case columns.
13. Rows whose Notes start with `EXAMPLE` are ignored by every parser.
14. `tools/qa_harness.js` is an in-page QA helper (`QA.sweep`, `QA.variant`, `QA.mutate` ...) used by the e2e tests.
15. Schemas: `tools/SCHEMA_KPI_WORKBOOK.md`, `tools/SCHEMA_POAP_WORKBOOK.md`, `docs/HOW_TO_FILL_DATA.md`.

## Commands
| Command | What it does |
|---|---|
| `npm ci` | install (only `@playwright/test`; Chromium is cached in `~/.cache/ms-playwright` / `~/Library/Caches/ms-playwright`, else `npx playwright install chromium`) |
| `npm run build` | `node --check` on every `assets/js/*.js` and `tools/*.js` (there is nothing to bundle) |
| `npm test` | unit tests, `node --test "tests/unit/*.test.js"` (node:test, no browser, a few seconds) |
| `npm run test:e2e` | Playwright, Chromium only; starts `tools/serve.js` itself; writes `playwright-report/` and `junit-qa.xml` |
| `npm run serve` | static server on http://localhost:8765 (`PORT` overrides), `Cache-Control: no-store` |
| `npm run qa` | build + test + test:e2e — **the gate** |

Node 22+ is required.

## How to add things
- **A KPI**: add an entry to `K.KPI_META` and a `kXxx(data, period, ctx)` function in `kpi-data.js`, register it in the `FN` map (return `empty(key, msg)` when there is no data, otherwise `base(key, p)` + `rag`, `position`, `metric`, `detail`, `highlights`, `concern`, `exec`). Add its visual in `kpi-pages.js` (`ATS.kpiPages[key] = (res, ctx) => ({ left, draw })`), list it in `monthly.js` and/or `weekly.js`, and if it needs new workbook columns extend `K.parse` (+ bump `ATS.VER`, regenerate sample data). Add unit tests in `tests/unit/`.
- **A page**: add `{ id, label, render: (env) => ({ html, draw }) }` to the `pages()` list of `monthly.js` / `weekly.js` (or `raid-pages.js` for RAID). Use `ATS.ui` (`U.slide`, `U.table`, `U.stat`) and `ATS.chart(key, canvas, config)` for charts. Make sure `QA.sweep` still reports nothing.
- **A RAID sheet**: parse it in `raid.js` (`RA.parse`; keep older logs without the sheet working, like Decisions), add it to `RA.summary` / `RA.quality`, a page in `raid-pages.js` (`ATS.raidPageList`), and unit tests.
- **An export**: PowerPoint slides live in `pptx-export.js` (`buildMonthlyOrWeekly`), PDF in `pdf-export.js`. Wire a button, trigger a real browser download (anchor + `a.download`), and add a `page.waitForEvent('download')` e2e check.

## Coding conventions
- Match the surrounding style (2-space indent, double quotes in the app files, `const`/arrow functions in newer files, `var` in `poap.js`). Keep changes small and local.
- **No new dependencies** (runtime or dev). No network calls at runtime. No `eval`, no inline event handlers.
- **Escape all user text** with `ATS.dom.esc` before putting it in HTML. Workbook text is untrusted.
- **Dates only via `ATS.date`** (`toIso`, `add`, `diff`, `networkdays` ...); **numbers only via `ATS.num`** (`num`, `ratio`, `pct`, ...). Never `new Date(string)` or `parseFloat` on cell values.
- Never let a bad cell crash a page: parsers return `null`/`[]` and push a message to the data-quality warnings instead.
- Do not reformat files you are not changing.

## Definition of done
1. `npm run qa` is green (build, unit, e2e). Do not weaken, skip or delete an existing test to get there; fix the code.
2. Every behaviour change has a test (unit for logic/parsing, Playwright for UI/export).
3. If the UI changed: open `index.html` (or `npm run serve`) in a browser and look at the affected page; check the browser console is clean.
4. Stay inside the diff budget (15 files / 600 changed lines) unless the work item says otherwise.
5. Thresholds in `qa-thresholds.json` are met (0 page errors, 0 sweep findings, >= 40 unit tests, e2e < 5 min).

## PROTECTED paths — never edit (human-owned)
`qa-thresholds.json` · `tests/quarantine.txt` · `.github/` · `pipelines/` · `azure-pipelines*` · `templates/*.xlsx` · `samples/**` · `assets/lib/**` · `assets/js/pptx-template.js` · `local_private/` · `ATS_Program/` · `AGENTS.md` · `.agent/` · `package.json` · `package-lock.json`

If a task seems to require changing one of these, stop and report it instead.

## Confidentiality
- The repo contains **fictional sample data only**. Never commit real programme data, real names, real TSR numbers or client wording anywhere (code, tests, fixtures, comments, commit messages).
- `ATS_Program/` and `local_private/` are git-ignored and must stay that way; never read them into test output or logs.
- Test data must be built in code from the fictional samples (load a sample workbook and mutate it) — do not commit new `.xlsx` files.

## How QA agents write tests
- Put new tests only under `tests/ai-generated/<workItemId>/`, as Playwright specs (`*.spec.js`). `playwright test` already picks them up. Do not edit `tests/e2e/`, `tests/unit/` or the config.
- Start the page like `tests/e2e/smoke.spec.js` does: `page.goto('/index.html')`, wait for `ATS.hasData()`, collect `pageerror` / `console.error`.
- Inject the helpers with `page.addScriptTag({ content: fs.readFileSync('tools/qa_harness.js', 'utf8') })`, then use `QA.sweep(portfolio)` (page-level defect scan), `QA.variant(...)` / `QA.mutate(...)` / `QA.toFile(...)` to craft edge-case workbooks in the browser from a sample, and `ATS.loadFiles([...])` to load them.
- Prefer real clicks and visible assertions (`getByRole`, `#view-title`, `.nav-item[data-view=...]`) over reaching into internals; downloads via `page.waitForEvent('download')`.
- A test must fail for the bug it targets and pass once fixed; keep each spec under ~60 s and the suite under the e2e time threshold.
- Report genuine bugs with the failing test, the observed vs expected behaviour and the work item id; never "fix" an app bug from the QA side.
