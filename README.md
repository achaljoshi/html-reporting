# ATS Program — Executive Dashboard

An offline, interactive reporting cockpit for the whole ATS programme. **Every portfolio fills the same weekly Excel workbook → you get each portfolio's Weekly Report, Monthly Council pack and POAP roadmap, plus a Programme Overview comparing all portfolios, plus ready-made PowerPoint.** No server, no internet, no screenshots, nothing uploaded anywhere.

| Section | What it is |
|---|---|
| **Programme Overview** | All portfolios side by side for a chosen month or week: RAG heat-map (portfolio × KPI), portfolio cards, "compare a KPI" ranking, and combined registers (TSRs, open defects, milestones at risk, high RAID items). Click anything to drill into that portfolio. |
| **Monthly Council** | The council deck as an interactive pack for the **selected portfolio** — scorecard, RAG history, RAID, monthly + quarterly KPI pages, commercials, resourcing, readiness. ← → keys page through it. |
| **Weekly Report** | The same portfolio data sliced by week: week-at-a-glance, what changed, burn-up, defects, environment, milestones, RAID, 12-week trends, look-ahead. |
| **POAP — Plan on a Page** | Interactive roadmap (Gantt, plan-on-a-page, integration flows) for the selected portfolio. |

The **Portfolio** dropdown (top bar) switches Monthly Council / Weekly Report / POAP between portfolios. The Programme Overview shows all of them at once.

---

## Start here
| I want to… | Go to |
|---|---|
| **Learn how to fill the data** (step by step, every column explained) | [`docs/HOW_TO_FILL_DATA.md`](docs/HOW_TO_FILL_DATA.md) — a Word copy to circulate to portfolio teams is in [`docs/ATS_Data_Filling_Guide.docx`](docs/ATS_Data_Filling_Guide.docx) |
| **See a finished, filled example** | `samples/ATS_Program_SAMPLE/Self Assessment/` (weekly workbook + POAP + RAID log — all fictional) and four more portfolios beside it |
| **Start my own portfolio** | copy `templates/ATS_Weekly_Data_Template.xlsx` (and `ATS_POAP_Template.xlsx`) — each has a built-in *Guide* sheet and an example row on every sheet |
| **Present** | open `index.html` |

## 1. Open it
Double-click **`index.html`** (Chrome or Edge). It opens with **sample data** — five fictional portfolios — and an amber *SAMPLE DATA* tag (bottom-right).

## 2. One design for every portfolio

**Each portfolio keeps its own copy of the same workbook** and fills it weekly. Nothing is merged by hand — the dashboard reads them all.

```
ATS_Program/                          ← the folder you load in the dashboard
├── Self Assessment/
│   └── ATS_Weekly_Data.xlsx          ← that portfolio's weekly workbook (copy of the template)
├── PAYE/
│   └── ATS_Weekly_Data.xlsx
├── Customs Declaration Service/
│   ├── ATS_Weekly_Data.xlsx
│   ├── ATS_POAP.xlsx                 ← optional: that portfolio's plan
│   └── ATS - Test RAID Log.xlsx      ← optional: that portfolio's RAID log (existing format, unchanged)
└── …one subfolder per portfolio
```
* **A portfolio = one weekly workbook.** Its name comes from **Config → Portfolio Name** (set it once; it must be unique).
* POAP and RAID are optional per portfolio. A portfolio without them simply shows "no POAP / RAID" on those pages.
* Prefer a flat folder? Name files `ATS_Weekly_Data - <Portfolio Name>.xlsx`; a POAP/RAID is matched to its portfolio by the portfolio name appearing in the file name (or inside the file).
* Blank templates left in the folder are ignored whenever real data exists.

### Weekly routine for each portfolio (≈5 minutes)
1. **Once:** copy `templates/ATS_Weekly_Data_Template.xlsx` into the portfolio's folder as `ATS_Weekly_Data.xlsx`; fill **Config** (Portfolio Name, budget, dates, thresholds); delete the grey *EXAMPLE* rows.
2. **Every Friday** add: one row in `Weekly_Snapshot`; one row per topic in `Topic_Progress`; add/update `Defect_Log`, `TSR_Log`, `Milestones` as things happen; `Readiness` (optional); `Commentary` (optional — your own wording).
3. **You (programme lead):** collect the folders into `ATS_Program/`, open the dashboard, click **Load program folder** (first time) or **Refresh data** (afterwards — Chrome/Edge remember the folder; click *Allow* once per session).
4. Use **Programme Overview** for the all-portfolio picture, or pick a portfolio and open **Weekly Report** / **Monthly Council** / **POAP**.
5. **Export PowerPoint** (Programme Overview, Monthly Council, Weekly Report) or **Copy as picture** on any page.

> **History matters:** every row is dated, so any week or month can be reopened later, trends build themselves, and week-over-week / month-over-month deltas are automatic.

### How a month is built from weeks
* **Flow measures are summed over the period:** planned test days, downtime hours.
* **Snapshot measures take the latest week in the period:** coverage, automation, test execution, resources, spend, forecast.
* **Event logs are filtered by date:** defects, TSRs, milestones (working days + UK bank holidays).
* Weeks belong to the month in which they **end**. Quarterly KPIs (leakage, CSAT, automation) show quarter-to-date.
* Everything is calculated **as of the end of the chosen week/month**, so last month's pack never changes when you add new data.

## 3. The workbook at a glance (`ATS_Weekly_Data.xlsx`)

| Sheet | Rows | Cadence | Feeds |
|---|---|---|---|
| Config | settings | once | portfolio name, thresholds, budget |
| Weekly_Snapshot | 1 / week | weekly | environment, coverage, automation, CSAT, resources, spend, overall RAG |
| Topic_Progress | 1 / topic / week | weekly | execution, burn-up, scope mix |
| Defect_Log | 1 / defect | as it happens | detection, rejection, aging/turnaround, leakage |
| TSR_Log | 1 / TSR | as it happens | TSR impact (working-day SLA) |
| Milestones | 1 / milestone | as it happens | milestone delivery |
| Readiness | 1 / topic / week | weekly | SIT readiness |
| Commentary | optional | weekly | your own highlights / concerns / actions |
| Resourcing, Demand_Forecast, Lessons_Learned | small tables | monthly | resourcing, demand pipeline, lessons |
| Bank_Holidays, Lists | reference | yearly | working-day maths, dropdowns |

Yellow = you type. Grey = formulas, don't type over. Rows whose **Notes** start with `EXAMPLE` are ignored. The workbook's **Guide** sheet explains every column.

### KPI rules (editable in `Config`)
| KPI | Rule |
|---|---|
| Risk Based Coverage | covered ÷ requirements per risk level; Green ≥ target, Amber within 2 pts, else Red; overall = worst of High & Medium |
| Production Defect Leakage (Q) | Red: P1 ≥ 1, P2 ≥ 5 or P3 ≥ 6; Amber: any P2/P3; *Not in Production* until go-live |
| Defect Detection Effectiveness | found in testing ÷ (testing + production); rejection % Green ≤ 5%, Amber ≤ 10% |
| Defect Aging | working-day buckets; SLA 5 days; Green ≥ 80%, Amber ≥ 60% *(assumption)* |
| Automation Coverage (Q) | Green if ≥ target or improving |
| Milestone Delivery | any Red milestone → Red; Amber if > 20% amber |
| Environment Availability | monthly downtime ≤ 24h Green, ≤ 48h Amber; weekly limits in Config *(assumption)* |
| TSR Impact | Red if any TSR > 10 working days (hold days excluded) |
| CSAT (Q) | pass mark 3; Green if ≥ pass and not declining |
| Resource / Demand / Commercial | plan-vs-actual, FTE gap, forecast-vs-budget *(assumptions)* |
| RAID Health | Red if a Very High item is open; Amber if High |
| Test Execution | blocked ÷ planned: Green ≤ 5%, Amber ≤ 20% *(assumption)* |

**RAG overrides:** an overall RAG typed in `Weekly_Snapshot` and any `RAG Override` in `Commentary` win over the computed colour — the pack shows both ("Declared Amber · computed Red").

## 4. Programme Overview
Pick **Monthly** or **Weekly** and a period; every portfolio is recalculated as of that date from its own workbook.
* **Heat-map** — rows = portfolios (worst first), columns = KPIs; click a dot to open that KPI page for that portfolio.
* **Portfolio cards** — overall RAG, Green/Amber/Red split, executed %, open defects, availability and the top concerns.
* **Compare a KPI** — rank all portfolios on one KPI (coloured by RAG); click a bar to drill in.
* **Registers** — TSRs, open defects, milestones at risk and High RAID items across all portfolios, filterable.
* The top strip totals open defects, TSRs over SLA, overdue milestones, spend vs budget and high RAID items for the whole programme.

## 5. POAP — Plan on a Page
Fill `ATS_POAP.xlsx` (copy of `templates/ATS_POAP_Template.xlsx`) once per portfolio and update dates/status as the plan moves.

| Sheet | One row per | Key columns |
|---|---|---|
| `POAP_Plan` | roadmap bar | Pillar · Topic (swim-lane) · Activity Type · Item · Start · End · **Baseline Start/End** · % Complete · Status · RAG · Owner · Depends On |
| `POAP_Milestones` | milestone | Pillar · Topic · Milestone · Date · Type · Status |
| `P2P_Matrix` | integration flow | Topic · System Path (`Portal → API Gateway → Ledger`) · Test No · Planned date · 2xx / 3xx / 4xx / 5xx results |
| `POAP_Config` | – | name, plan window, **Status Date** (blank = today), colours |

Three views, all filterable: **Roadmap** (swim-lanes, % fill, RAG, baseline ghost bar + slip badge, milestones, today line, minimap, zoom, dependency arrows, detail drawer), **Plan on a Page** (one-screen poster of phase chevrons per topic, next-30-days and at-risk panels) and **Integration Flows** (system chips; click a system to see every flow touching it). The POAP also feeds the Weekly Report look-ahead.

> The sample POAP (Self Assessment) is fictional. Two bars carry an illustrative baseline slip ("SAMPLE slip for demo") and a couple are `Dates TBC` — they show what those states look like; clear or replace them in your own file.

## 6. Accuracy features
* TSR days, defect aging and milestone delays use **working days with UK bank holidays**.
* **Data Quality page** (Appendix of Monthly/Weekly) flags closed defects without a resolved date, returned-before-received TSRs, non-Friday week endings, a stale weekly row, etc.
* Auto-drafted commentary is labelled **Auto-drafted**; your own `Commentary` text is labelled **From Excel**.
* A RAID log's own summary (open risks by rating band, open dependencies) is reproduced exactly.

## 7. Folder layout of this project
```
ATS_Dashboard/
├── index.html                      ← open this
├── docs/                           ← HOW_TO_FILL_DATA.md + ATS_Data_Filling_Guide.docx
├── templates/                      ← blank templates to copy for each portfolio team
│   ├── ATS_Weekly_Data_Template.xlsx
│   └── ATS_POAP_Template.xlsx
├── samples/ATS_Program_SAMPLE/     ← five fictional portfolios used as the built-in demo
│   └── Self Assessment/            ←   (also has a sample POAP + RAID log)
├── ATS_Program/                    ← put YOUR portfolio folders here (see section 2; git-ignored)
├── assets/                         ← code + bundled libraries (SheetJS, Chart.js, PptxGenJS, html2canvas) — all offline
├── tools/                          ← schema docs, sample builder, guide builder, QA server
├── presentation/                   ← the intro deck
└── archive/legacy_portfolio_explorer/   ← the first multi-portfolio explorer (superseded)
```

## 8. Confidentiality
Everything committed to this repository is **fictional sample data** and blank templates. Your real material must never be committed: `ATS_Program/` (your live portfolio folders) and `local_private/` are in `.gitignore`. Generators that embed real programme wording stay local-only (also git-ignored). If you add real data anywhere else, check it before pushing.

## 9. Troubleshooting
| Symptom | Fix |
|---|---|
| A portfolio is missing from the dropdown | its workbook needs the sheets `Weekly_Snapshot` + `Defect_Log` and at least one non-EXAMPLE row |
| Two portfolios merged / one disappeared | `Portfolio Name` in Config must be unique per workbook |
| "Workbook is loaded but has no weekly rows yet" | add a row to `Weekly_Snapshot` (delete the EXAMPLE row) and click Refresh |
| POAP / RAID not shown for a portfolio | put the file in that portfolio's folder (or include the portfolio name in the file name) |
| A KPI shows grey / "No data" | the sheet behind it has no rows up to that period |
| Refresh asks for the folder again | normal in Firefox/Safari; Chrome/Edge ask to *Allow* once per session |
| Export PowerPoint does nothing | the browser blocked the download — allow downloads for the page |

## 10. Customising
* Thresholds, budget, dates → `Config`. Topics / dropdown values → `Lists`.
* Colours/branding → CSS variables at the top of `assets/css/style.css`.
* Rebuild the embedded demo after changing files in `samples/`: `node tools/build_sample_js.js`.
* Rebuild the Word/Markdown filling guide after editing its content: `NODE_PATH=$(npm root -g) node tools/build_filling_guide.js`.
* Workbook layouts are specified in `tools/SCHEMA_KPI_WORKBOOK.md` and `tools/SCHEMA_POAP_WORKBOOK.md`.
