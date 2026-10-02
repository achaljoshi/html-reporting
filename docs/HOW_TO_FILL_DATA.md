# ATS Dashboard — How to Fill the Data

*Step-by-step guide for portfolio teams and the programme lead*


## 1. The idea in 60 seconds

- **One workbook per portfolio.** Every portfolio fills the *same* Excel workbook (`ATS_Weekly_Data.xlsx`, copied from `templates/ATS_Weekly_Data_Template.xlsx`).
- **Fill it weekly, never overwrite history.** Each Friday you add a new dated row (and update any defects, TSRs or milestones that changed). Old rows stay — that is what builds the trends.
- **Nothing is entered twice.** The Weekly Report, the Monthly Council pack, the Programme Overview and the PowerPoint export are all calculated from the same rows. A month is simply the weeks that end in it.
- **Optional extras per portfolio:** a POAP workbook (the plan on a page) and a RAID log (your existing RAID file, unchanged).

> ℹ️ Look at a finished example first: open `samples/ATS_Program_SAMPLE/Self Assessment/ATS_Weekly_Data.xlsx` (and its `ATS_POAP.xlsx` and `ATS_RAID_Log.xlsx`) next to this guide. All sample data is fictional.


## 2. What you get in the box

| File / folder | What it is | Who uses it |
|---|---|---|
| `templates/ATS_Weekly_Data_Template.xlsx` | Blank weekly workbook with an example row on every sheet and a built-in **Guide** sheet. | Every portfolio team |
| `templates/ATS_POAP_Template.xlsx` | Blank plan-on-a-page workbook (plan bars, milestones, integration flows). | Portfolios that want a roadmap |
| `samples/ATS_Program_SAMPLE/` | Five fictional portfolios (Self Assessment, PAYE, VAT, Customs Declaration Service, Child Benefit). Self Assessment also has a POAP and a RAID log. | Everyone — to see how a filled workbook looks |
| `ATS_Program/` | Your real data folder. Put one subfolder per portfolio here. | Programme lead |
| `index.html` | The dashboard. Double-click to open (Chrome or Edge). | Everyone |


## 3. One-time set-up for each portfolio

1. Create a folder for the portfolio inside `ATS_Program/`, e.g. `ATS_Program/Self Assessment/`.
2. Copy `templates/ATS_Weekly_Data_Template.xlsx` into it and rename it `ATS_Weekly_Data.xlsx`.
3. Open the **Config** sheet. Set **Portfolio Name** (must be unique across portfolios — it is the name in the dashboard dropdown), **Programme**, **Program Name**, **Project / Phase**, **TSR Reference**, the managers, **Total Budget (£)** and the programme dates. Set **Data Mode** to `LIVE`.
4. Review the thresholds further down Config (SLA days, downtime limits, coverage targets). The defaults are sensible; change them only if your contract says otherwise.
5. On the **Lists** sheet, replace the example **Topic** values with your portfolio's own topics / releases / workstreams (they feed the dropdowns).
6. **Delete the grey EXAMPLE rows** on each data sheet (the row whose Notes column says "EXAMPLE – delete this row"). Rows marked EXAMPLE are ignored by the dashboard, but delete them anyway.
7. Optional: copy `templates/ATS_POAP_Template.xlsx` as `ATS_POAP.xlsx` and your RAID log into the same folder.

```
ATS_Program/
├── Self Assessment/
│   ├── ATS_Weekly_Data.xlsx      <- required
│   ├── ATS_POAP.xlsx             <- optional
│   └── ATS_RAID_Log.xlsx         <- optional (your existing RAID format)
├── PAYE/
│   └── ATS_Weekly_Data.xlsx
└── VAT/
    └── ATS_Weekly_Data.xlsx
```

> ℹ️ Prefer one flat folder? Name the files `ATS_Weekly_Data - <Portfolio Name>.xlsx`. A POAP or RAID file is matched to its portfolio when the portfolio name appears in its file name.


## 4. The weekly routine (every Friday, about 5 minutes)

| Sheet | What to do | How many rows |
|---|---|---|
| **Weekly_Snapshot** | Add ONE new row for this week (Week Ending = the Friday). Enter downtime, coverage, automation, resources, spend and your overall RAG with a one-paragraph summary. | 1 per week |
| **Topic_Progress** | Add one row per active topic with cumulative test-case counts to date. | 1 per topic per week |
| **Defect_Log** | Add new defects; update status / Resolved Date of existing ones. Never delete a defect. | as it happens |
| **TSR_Log** | Add new TSRs when received; fill Returned Date and Hold Days when you respond. | as it happens |
| **Milestones** | Fill the Completed Date when a milestone is done; adjust the Forecast Date if it moves. | as it happens |
| **Readiness** | Add one row per topic with six RAG areas and the Road to Green. | 1 per topic per week (optional) |
| **Commentary** | Optional: your own wording for any KPI. If you leave it empty the dashboard drafts the text for you. | optional |

**Monthly / quarterly items** (when they change): CSAT score in `Weekly_Snapshot`, `Demand_Forecast`, `Resourcing`, `Lessons_Learned`.

Then save the workbook. The programme lead opens the dashboard and clicks **Refresh data** (or **Load program folder** the first time).

> ⚠️ **Golden rules:** (1) week ending dates must be **Fridays**; (2) yellow cells are yours, grey cells are formulas — never type over grey; (3) counts in Weekly_Snapshot and Topic_Progress are **cumulative to date**, not "this week only"; (4) never delete history; (5) keep every portfolio's **Portfolio Name** unique.


## 5. Sheet-by-sheet reference

Yellow = you type. Grey = calculated. The header text must stay exactly as in the template — the dashboard finds columns by their header.


### 5.1 Config — settings (once)

| Setting | What to enter | Example |
|---|---|---|
| Portfolio Name | The name shown in the dashboard. Unique per portfolio. | Self Assessment |
| Programme | Groups portfolios under one heading. | HMRC ATS |
| Program Name | Full title used in page headers. | HMRC ATS – Self Assessment |
| Project / Phase, TSR Reference | Shown on commercial pages. | Phase 1 / TSR_PR021877-REQ031204 |
| Programme Test Manager, Test Manager | Names for the resourcing page. | role or name |
| Total Budget (£) | Approved budget; the commercial KPI compares the forecast with it. | 760000 |
| Production Go-Live Date | Leave blank until live. While blank, "Production Defect Leakage" shows *Not in Production*. | (blank) |
| Data Mode | `LIVE` for real data, `SAMPLE` for demo data (adds a SAMPLE tag in the dashboard). | LIVE |
| Thresholds | TSR SLA (10 working days), defect SLA (5), rejection %, downtime limits, coverage targets, milestone window, automation target, CSAT pass mark, execution blocked limits. | defaults provided |


### 5.2 Weekly_Snapshot — one row per week

| Column(s) | What to enter | Example (Self Assessment, W/E 2 Oct 2026) |
|---|---|---|
| Week Ending | The Friday of the reporting week. | 2026-10-02 |
| Planned Test Days | Working days testing was planned that week (bank-holiday weeks: 4). | 5 |
| Downtime Hours | Hours test execution was suspended (environment outage etc.). **Availability %** is calculated. | 0 |
| High / Medium / Low Risk – Requirements | Total requirements tagged at that risk level (cumulative). | 72 / 100 / 48 |
| … – Test Cases | Test cases written against that risk level. | 126 / 189 / 105 |
| … – Covered Reqs | Requirements fully covered by scripts. **Coverage %** is calculated (covered ÷ requirements). | 72 / 97 / 30 |
| Automatable / Automated Test Data Setups | How many test-data set-ups can be automated, and how many are. | 150 / 108 |
| Automatable / Automated Test Cases | Same for test-case execution automation. | 200 / 150 |
| CSAT Score (1-5), CSAT Responses | Only in a week a score was collected; otherwise leave blank ("Not requested"). | (blank) |
| Resources Planned / Actual (FTE) | Planned vs actual full-time equivalents in place. | 6 / 6 |
| Cumulative Spend (£), Forecast at Completion (£) | Spend to date and the latest forecast. | 608000 / 760000 |
| Overall RAG | Your judgement: Green, Amber or Red. Shown alongside the RAG the dashboard calculates. | Green |
| Overall Executive Summary | One paragraph for the council / report. | "Portfolio remains Green: execution on plan…" |


### 5.3 Topic_Progress — one row per topic per week (cumulative)

| Column | What to enter | Example |
|---|---|---|
| Week Ending, Topic | The Friday and the topic / release / workstream (pick from the dropdown). | 2026-10-02, Self Assessment – Digital Filing |
| Phase | Test Prep, Smoke, SIT, Regression, E2E, NFR or TCR / Sign-off. | SIT |
| TCs Planned / Designed | Test cases planned in scope, and designed (scripted) so far. | 70 / 63 |
| Passed / Failed / Blocked | Cumulative results to date. **Executed**, **% Executed** and **% Designed** are calculated. | 4 / 0 / 0 |
| Comments | Short note. | 4 executed (4 passed, 0 failed); 0 blocked. |


### 5.4 Defect_Log — one row per defect

| Column | What to enter | Example |
|---|---|---|
| Defect ID, Summary | From your defect tool (Jira etc.). | SA-101, Tax calculation – message queued but not delivered |
| Topic, Test Type | Topic dropdown; Smoke, SIT, E2E, NFR, Regression or Production. | Self Assessment – Release 1, Smoke |
| Priority | P1 High, P2 Medium or P3 Low. | P2 Medium |
| Severity | 1 Urgent … 6 Trivial (see the definitions in the Guide). | 3 - High |
| Raised Date | Date raised. | 2026-08-05 |
| Resolved Date | The date work **stops** (fix and confirmation test done, or the date it was rejected). Blank while open. | 2026-08-06 |
| Status | Open, In Progress, Re-Open, Resolved, Closed, Rejected or Deferred. | Closed |
| Found In | Testing or Production. Production defects drive the leakage KPI. | Testing |
| Owner / DG, TCs Blocked, Triage Comments | Who is fixing it, test cases it blocks, latest note. | Data Team, 0, Fix delivered and confirmed. |
| Aging, Aging Bucket (grey) | Calculated in **working days**, bank holidays excluded: 0-2, 3-5, 6-10, >10 days. | 2, 0-2 Days |

> ℹ️ A **Closed** or **Resolved** defect must have a Resolved Date — otherwise its aging cannot be measured (the Data Quality page will flag it).


### 5.5 TSR_Log — one row per Test Scope Request

| Column | What to enter | Example |
|---|---|---|
| TSR Ref, Description, Topic | Reference, short description, topic. | PR022310-REQ032588 |
| Received Date | Date the TSR reached you. | 2026-09-28 |
| Returned Date | Date you sent it back to the front desk with commercials complete. Blank while open (it then ages to today). | (blank) |
| Hold Days (pending clarification) | Working days you were waiting for clarification — these are subtracted. | 0 |
| Status | Open, Returned or Withdrawn (withdrawn TSRs are ignored). | Open |
| Work Days, Over SLA? (grey) | Calculated: working days minus hold days; "Yes" if above the limit in Config (default 10). | 5, No |


### 5.6 Milestones — one row per milestone

| Column | What to enter | Example |
|---|---|---|
| Milestone, Topic, Owner | Name, topic (or Programme), owner role. | Test Completion Report, Programme (cross-cutting) |
| Due Date | The agreed (baseline) date. | 2026-12-18 |
| Forecast Date | Optional: the current forecast if it has moved. | (blank) |
| Completed Date | Fill when done. Completed after the Due Date = late. | (blank) |
| Critical? | Yes for milestones that count towards "critical milestone delay". | No |
| Delay, Status, RAG (grey) | Calculated: Completed / Completed Late / Overdue / On Schedule; Amber in the last 5 days before the due date, Red once overdue. | 0, On Schedule, Green |


### 5.7 Readiness — one row per topic per week (optional)

Columns: **Week Ending, Topic, Planned SIT Start, Overall RAG**, then for each of six areas — *Requirements & Arch Intent, Blockers, Resources, Environments, Test Scripts & Test Data, Knowledge Transfer* — a **RAG** and a **Note**, and finally **Road to Green** (what has to happen to turn the topic Green). Example: Overall Green, every area Green, note "Requirements baselined and reviewed."


### 5.8 Commentary — your own words (optional)

One row per KPI per week: **Week Ending, KPI** (pick from the list), **Key Highlights, Area of Concern, Actions Underway, Executive Commentary**, optional **RAG Override** and **Current Position Override**. Put each bullet on its own line (Alt+Enter). Anything you type here replaces the auto-drafted text for that KPI and is labelled *From Excel* in the dashboard; blank fields keep the auto draft.


### 5.9 Small tables

| Sheet | Columns | Notes |
|---|---|---|
| Resourcing | Role, Name, Organisation, Allocation %, Start Date, End Date, Status | Status: Active, Planned or Left. Use role names if you prefer not to list people. |
| Demand_Forecast | Month (first day), Planned FTE, Available FTE, Notes | A gap of more than 1 FTE in the next 3 months turns the Demand KPI Red. |
| Lessons_Learned | Date, Category, Lesson Learned, Improvement Action, Owner, Status | Shown on the "Lessons Learned" council page. |
| Bank_Holidays | Date, Name | Used for all working-day maths. Add next year's dates each January. |
| Lists | Dropdown values | Edit Topic here; leave the rest unless you add a new value. |


## 6. POAP — plan on a page

Fill `ATS_POAP.xlsx` once and update dates and status as the plan moves. The dashboard turns it into a roadmap (Gantt), a one-page poster and an integration-flow explorer.

| Sheet | One row per | Key columns |
|---|---|---|
| POAP_Config | setting | Program Name, Plan Start/End Date, **Status Date** (blank = today), Data Mode, Pillars and Activity Types with colours |
| POAP_Plan | bar on the roadmap | ID, Pillar, Topic (swim-lane), Activity Type, Item, **Start, End**, **Baseline Start, Baseline End**, % Complete, Status, RAG, Owner, Depends On, Scope / Notes |
| POAP_Milestones | milestone | ID, Pillar, Topic, Milestone, Date, Type (SIT Start, SIT End, Sign-off / TCR, Go-Live, Gate…), Status (Planned, Achieved, At Risk, Missed) |
| P2P_Matrix | integration flow | Flow ID, Topic, Sub Process, **System Path** (`System A → System B → System C`), Test No, Planned Date, Status, results for 2xx / 3xx / 4xx / 5xx |


#### How to fill POAP_Plan

- **Start / End** are the *current* plan. **Baseline Start / End** are the *originally agreed* dates — fill them once and never change them. The dashboard shows the gap as a ghost bar and a slip badge such as **+14d** ("Slip (Days)" is calculated).
- **Status:** Not Started, In Progress, Complete, Delayed, On Hold or Dates TBC. For *Dates TBC* leave Start and End blank.
- **% Complete** is 0–100% (type 0.5 or 50%).
- **Depends On:** the IDs this bar waits for, comma-separated (e.g. `A-003, A-004`). The roadmap draws the dependency arrows.
- **Activity Type** decides the bar colour. Use the dropdown: Test Prep, Dependencies, Regression, SIT, Smoke / P2P, Defect Retest, TCR / Sign-off, E2E, NFR, Automation, Holiday / Freeze, Other.
- Rows whose Notes start with `EXAMPLE` are ignored — delete the example row once you have real rows.


## 7. RAID log

Keep using your existing RAID log exactly as it is — no changes. Put it in the portfolio's folder. The dashboard reads the **Risks, Issues, Assumptions and Dependencies** sheets by their header names and shows open items by rating (Very Low / Low / Medium / High / Very High), a likelihood × impact heat-map and the items needing attention. **RAID Health** is Red when a Very High item is open and Amber when a High item is open. A closed risk or a resolved issue drops out automatically.


## 8. After you fill: load and check

1. Open `index.html`. Click **Load program folder** and choose `ATS_Program` (first time), or **Refresh data** after later edits.
2. Use the **Portfolio** dropdown (top bar) to switch between portfolios, or **Programme Overview** to see them all.
3. Open **Appendix → Data Quality** in the Monthly Council or Weekly Report. It lists problems such as a closed defect without a Resolved Date, a TSR returned before it was received, or a week ending that is not a Friday.
4. Use **Export PowerPoint**, or **Copy as picture** on any page, for your slides.


## 9. How the numbers are worked out

| KPI | Rule (thresholds are editable in Config) |
|---|---|
| Risk Based Coverage | Covered ÷ requirements per risk level. Green at or above target; Amber within 2 points; else Red. Overall = worst of High and Medium. |
| Production Defect Leakage (quarterly) | Red if any P1, or 5+ P2, or 6+ P3 found in production in the quarter; Amber for any P2/P3; else Green. |
| Defect Detection Effectiveness | Defects found in testing ÷ (testing + production). Rejection rate Green ≤ 5%, Amber ≤ 10%. |
| Defect Aging | Working-day buckets; Green if at least 80% resolved within 5 working days, Amber at least 60%, else Red. |
| Automation Coverage (quarterly) | Green if at or above target, or improving on the previous period. |
| Milestone Delivery | Red if any milestone is overdue or late; Amber if more than 20% are Amber. |
| Environment Availability | 1 − downtime ÷ (test days × 8). Monthly: Green ≤ 24h, Amber ≤ 48h. Weekly limits are in Config. |
| TSR Impact Assessment | Red if any TSR takes more than 10 working days (hold days excluded). Open TSRs age to today. |
| Customer Satisfaction (quarterly) | Pass mark 3 of 5; Green if at or above it and not declining. |
| Resource / Demand / Commercial | Actual ÷ planned FTE; FTE gap in the next 3 months; forecast vs budget (Amber within 5% over). |
| RAID Health | Red with a Very High item open; Amber with a High item open. |
| Test Execution Progress | Blocked ÷ planned: Green ≤ 5%, Amber ≤ 20%, else Red. |


#### How weekly rows become a month

- **Added up:** planned test days and downtime hours.
- **Latest week in the period:** coverage, automation, test execution, resources, spend, forecast.
- **Filtered by date:** defects, TSRs, milestones.
- A week belongs to the month it **ends** in. Every month is calculated as of its last day, so last month's pack never changes when you add new rows.


## 10. Common mistakes

| Mistake | What happens | Fix |
|---|---|---|
| Week Ending is not a Friday | Flagged on the Data Quality page; the week may land in the wrong month. | Use the Friday of that week. |
| Typed "this week only" numbers in Weekly_Snapshot / Topic_Progress | Coverage and execution fall instead of grow. | Enter cumulative totals to date. |
| Deleted old defect/TSR rows | Past months change and trends break. | Keep history; change Status instead. |
| Closed a defect without a Resolved Date | Aging cannot be measured. | Add the date work stopped. |
| Two workbooks with the same Portfolio Name | One portfolio disappears from the dropdown. | Make the names unique. |
| Typed over a grey formula cell | Calculated column stops updating. | Restore the formula from the template. |
| Left the EXAMPLE row in | Ignored by the dashboard, but confusing. | Delete it. |
| Forgot to click Refresh data | The dashboard still shows the previous load. | Click Refresh data. |


## 11. Glossary

| Term | Meaning |
|---|---|
| RAG | Red / Amber / Green status. |
| TSR | Test Scope Request — a request to estimate or scope testing, answered within an SLA. |
| SLA | The agreed response time (default 10 working days for a TSR, 5 for defect turnaround). |
| SIT / E2E / NFR | System Integration Test, End-to-End test, Non-Functional test. |
| P2P | Point-to-point integration test between two systems. |
| POAP | Plan on a Page — the programme roadmap. |
| RAID | Risks, Assumptions, Issues, Dependencies. |
| FTE | Full-time equivalent. |
