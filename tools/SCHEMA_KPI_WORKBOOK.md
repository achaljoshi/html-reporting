# ATS Weekly Data Workbook — schema (single source of truth)

One workbook, filled **weekly**, feeds BOTH the Weekly Report and the Monthly Council.
Everything is stored as **dated rows (history is never overwritten)**. A "month" is a roll-up of weeks/events; a "week" is a slice.

File names (templates): `ATS_Weekly_Data_Template.xlsx` (blank + example rows) and, per sample portfolio, `samples/ATS_Program_SAMPLE/<Portfolio Name>/ATS_Weekly_Data.xlsx` (fictional seeded demo workbooks).

**One design for every portfolio:** each portfolio keeps its own copy of the workbook and identifies itself with `Config > Portfolio Name` (unique across portfolios) and `Config > Programme` (groups portfolios). Recommended layout `ATS_Program/<Portfolio Name>/ATS_Weekly_Data.xlsx` (+ that portfolio's `ATS_POAP.xlsx` and RAID log); alternative flat folder with `ATS_Weekly_Data - <Portfolio Name>.xlsx`.

## Global conventions (every data sheet)
* Row 1 = sheet title (bold, 14pt, navy `1B2A4A`). Row 2 = one-line purpose. Row 3 = "HOW TO FILL" one-liner. Row 4 = blank.
* **Row 5 = header row. Row 6 onward = data.** (Dashboard parser finds columns by header text, case-insensitive, so keep header text EXACTLY as below.)
* Blank template: row 6 holds ONE realistic **example row** (grey italic font). The last column of EVERY data sheet is `Notes`; the example row has `Notes` = `EXAMPLE – delete this row`. **The parser skips any row whose Notes starts with "EXAMPLE".**
* Input cells: yellow fill `FFF6CC`, blue font `0000FF`. Calculated cells ("calc"): light grey fill `F2F5F9`, black font, never typed over. Header: navy fill `1B2A4A`, white bold. Font Arial 10/11 throughout.
* Dates are real Excel dates formatted `yyyy-mm-dd`. Freeze panes below header (and first column where sensible). Autofilter on header row. Pre-format ~300 data rows (fill, validation, calc formulas) in the blank template.
* Calc columns use Excel formulas (Excel-2007-era functions only: NETWORKDAYS, IF, IFERROR, COUNTIF, SUMIFS, INDEX/MATCH). Run LibreOffice recalc afterwards (script: `…/skills/xlsx/scripts/recalc.py`) so cached values exist; must finish with 0 errors. Formulas must tolerate blank rows (return "").
* Dropdown (list) data validation on every enumerated column, referencing the `Lists` sheet.
* Sample/Live flag: `Config!Data Mode` = `SAMPLE` or `LIVE`. SAMPLE workbook must say SAMPLE; blank template says `LIVE`.

## Sheet list (exact names, in this order)
1. `Guide`  2. `Config`  3. `Weekly_Snapshot`  4. `Topic_Progress`  5. `Defect_Log`  6. `TSR_Log`  7. `Milestones`  8. `Readiness`  9. `Commentary`  10. `Resourcing`  11. `Demand_Forecast`  12. `Lessons_Learned`  13. `Bank_Holidays`  14. `Lists`

### 1. Guide
Plain-English instructions. Must contain: (a) the "fill once, weekly" idea; (b) a table "What to update and when": Weekly = Weekly_Snapshot (1 new row), Topic_Progress (1 row per active topic), Defect_Log (add/update defects), Readiness (1 row per topic), Commentary (optional narrative for the week); As-they-happen = TSR_Log, Milestones, Defect_Log; Monthly/Quarterly = CSAT score (when collected), Demand_Forecast, Lessons_Learned, Resourcing; (c) colour legend; (d) how flow vs cumulative measures roll up: **Flow measures (SUM over the period): Planned Test Days, Downtime Hours. Cumulative/snapshot measures (take the LATEST week in the period): coverage, automation, test execution, resources, spend, forecast.** (e) the RAG rules in a table (see Rules below); (f) the rule "rows whose Notes start with EXAMPLE are ignored"; (g) the RAID log is maintained in its own existing file (`ATS - Test RAID Log.xlsx` format) and the POAP in `ATS_POAP.xlsx` — drop all three files in the same folder and load that folder in the dashboard.

### 2. Config  (header row 4: `Setting | Value | Description / assumption`, data from row 5; key in col A, value in col B)
Exact keys (parser reads col A text):
```
Program Name                                  HMRC ATS – Self Assessment
Project / Phase                               Phase 1
TSR Reference                                 TSR_PR021877-REQ031204
Programme Test Manager                        TBC
Test Manager                                  TBC
Programme Start Date                          2026-07-01
Programme End Date                            2028-05-29
Total Budget (£)                              760000
Production Go-Live Date                       (blank = not yet in production)
Data Mode                                     SAMPLE | LIVE
Hours per Test Day                            8
TSR SLA – Max Working Days                    10
Defect Turnaround SLA – Working Days          5
Defect Turnaround – Green if SLA compliance >=   0.80
Defect Turnaround – Amber if SLA compliance >=   0.60
Defect Rejection – Green if <=                0.05
Defect Rejection – Amber if <=                0.10
Downtime (monthly) – Green if hours <=        24
Downtime (monthly) – Amber if hours <=        48
Downtime (weekly) – Green if hours <=         6
Downtime (weekly) – Amber if hours <=         12
Coverage Target – High Risk                   0.98
Coverage Target – Medium Risk                 0.95
Coverage Target – Low Risk                    0
Coverage – Amber tolerance (points)           0.02
Milestone – Amber window (days before due)    5
Milestone – Amber if share of Amber >         0.20
Automation Target                             0.95
CSAT Pass Mark (1-5 score)                    3
Resource – Amber if actual >= (share of plan) 0.90
Execution – Amber if blocked share of planned >  0.05
Execution – Red if blocked share of planned >    0.20
Commercial – Amber if forecast <= budget + (share)  0.05
Portfolio Name                                Self Assessment   (template: <Portfolio Name>; unique across portfolios; shown in the dashboard's portfolio dropdown and cross-portfolio overview)
Programme                                     HMRC ATS       (groups portfolios under one programme heading)
```
The two portfolio keys are APPENDED after all existing keys (Config rows B38 and B39 at the time of writing; existing rows never move). Execution keys: blocked ÷ planned test cases, ASSUMPTION. `Program Name` is unchanged (e.g. `HMRC ATS – Self Assessment`).
Column C documents the source of every rule (e.g. "from ATS KPI Data Collation – M.TSR Assessment (3.5)") and flags **assumptions** (weekly downtime thresholds, turnaround RAG thresholds, resource/commercial thresholds are ASSUMPTIONS – say so).

### 3. Weekly_Snapshot  — ONE ROW PER WEEK (header row 5)
Group band in row 4 (merged, coloured) over the headers: ENVIRONMENT | COVERAGE | AUTOMATION | SATISFACTION | RESOURCING & COMMERCIALS | OVERALL.
```
A  Week Ending                      date (Friday). input
B  Planned Test Days                number (default 5). input
C  Downtime Hours                   number. input
D  Availability %                   calc = IF(B="","",1-C/(B*Config!HoursPerTestDay))
E  High Risk – Requirements         input
F  High Risk – Test Cases           input
G  High Risk – Covered Reqs         input
H  High Risk – Coverage %           calc = IF(E>0,G/E,"")
I  Medium Risk – Requirements       input
J  Medium Risk – Test Cases         input
K  Medium Risk – Covered Reqs       input
L  Medium Risk – Coverage %         calc
M  Low Risk – Requirements          input
N  Low Risk – Test Cases            input
O  Low Risk – Covered Reqs          input
P  Low Risk – Coverage %            calc
Q  Automatable Test Data Setups     input
R  Automated Test Data Setups       input
S  Test Data Automation %           calc
T  Automatable Test Cases           input
U  Automated Test Cases             input
V  Test Case Automation %           calc
W  CSAT Score (1-5)                 input (leave blank unless a score was collected that week)
X  CSAT Responses                   input
Y  Resources Planned (FTE)          input
Z  Resources Actual (FTE)           input
AA Cumulative Spend (£)             input
AB Forecast at Completion (£)       input
AC Overall RAG                      input dropdown Green/Amber/Red
AD Overall Executive Summary        input (long text, wrap)
AE Notes
```
Coverage / execution / automation / resources / spend values are **cumulative to date** at that week-ending.

### 4. Topic_Progress — ONE ROW PER TOPIC PER WEEK (header row 5; all counts cumulative to date)
```
A Week Ending | B Topic | C Phase | D TCs Planned | E TCs Designed | F Passed | G Failed | H Blocked |
I Executed (calc =F+G) | J % Executed (calc =IF(D>0,I/D,"")) | K % Designed (calc =IF(D>0,E/D,"")) | L Comments | M Notes
```
Phase list: `Test Prep`, `Smoke`, `SIT`, `Regression`, `E2E`, `NFR`, `TCR / Sign-off`.

### 5. Defect_Log — ONE ROW PER DEFECT (header row 5)
```
A Defect ID | B Summary | C Topic | D Test Type | E Priority | F Severity | G Raised Date | H Resolved Date |
I Status | J Found In | K Owner / DG | L TCs Blocked | M Triage Comments | N Aging (Working Days) calc | O Aging Bucket calc | P Notes
```
* Priority list: `P1 High`, `P2 Medium`, `P3 Low`. Severity list: `1 - Urgent`,`2 - Very High`,`3 - High`,`4 - Standard`,`5 - Low`,`6 - Trivial`.
* Test Type list: `Smoke`,`SIT`,`E2E`,`NFR`,`Regression`,`Production`. Status list: `Open`,`In Progress`,`Re-Open`,`Resolved`,`Closed`,`Rejected`,`Deferred`.
* Found In list: `Testing`, `Production`.
* `Resolved Date` = date work stops on the defect (fix + confirmation test complete, or rejection date).
* N = NETWORKDAYS(G, IF(H="",TODAY(),H), Bank_Holidays dates) (inclusive, Excel semantics; min 1 when raised). O bucket: `0-2 Days`, `3-5 Days`, `6-10 Days`, `>10 Days`.

### 6. TSR_Log — ONE ROW PER TSR (header row 5)
```
A TSR Ref | B Description | C Topic | D Received Date | E Returned Date | F Hold Days (pending clarification) |
G Status | H Value (£) | I Work Days calc | J Over SLA? calc | K Notes
```
Status list: `Open`,`Returned`,`Withdrawn`. I = MAX(0, NETWORKDAYS(D, IF(E="",TODAY(),E), holidays) - F). J = IF(I>Config SLA,"Yes","No").

### 7. Milestones (header row 5)
```
A Milestone | B Topic | C Due Date | D Forecast Date | E Completed Date | F Owner | G Critical? | H Delay (Days) calc | I Status calc | J RAG calc | K Notes
```
Critical? list `Yes`,`No`. H = IF(E<>"",MAX(0,E-C),MAX(0,TODAY()-C)). I = `Completed` / `Completed Late` / `Overdue` / `On Schedule`. J (from ATS KPI Data Collation – M.Test Milestone Delays): if completed: Green if E<=C else Red; else if TODAY()>C Red; else if C-TODAY()<=Config amber window Amber; else Green.

### 8. Readiness — ONE ROW PER TOPIC PER WEEK (header row 5)
```
A Week Ending | B Topic | C Planned SIT Start | D Overall RAG |
E Requirements & Arch Intent RAG | F Requirements & Arch Intent – Note |
G Blockers RAG | H Blockers – Note | I Resources RAG | J Resources – Note |
K Environments RAG | L Environments – Note | M Test Scripts & Test Data RAG | N Test Scripts & Test Data – Note |
O Knowledge Transfer RAG | P Knowledge Transfer – Note | Q Road to Green | R Notes
```
### 9. Commentary — optional narrative per KPI per week (header row 5)
```
A Week Ending | B KPI | C Key Highlights | D Area of Concern | E Actions Underway | F Executive Commentary | G RAG Override | H Current Position Override | I Notes
```
KPI list (exact): `Overall`,`Risk Based Coverage`,`Production Defect Leakage`,`Defect Detection Effectiveness`,`Defect Aging`,`Automation Coverage`,`Milestone Delivery`,`Environment Availability`,`TSR Impact Assessment`,`Customer Satisfaction (CSAT)`,`Resource Position`,`Demand Forecast`,`Commercial Performance`,`RAID Health`,`Test Execution Progress`,`SIT Readiness`.
Multi-line text: separate bullets with line breaks (Alt+Enter).

### 10. Resourcing (header row 5): `Role | Name | Organisation | Allocation % | Start Date | End Date | Status | Notes`  (Status list `Active`,`Planned`,`Left`)
### 11. Demand_Forecast (header row 5): `Month | Planned FTE | Available FTE | Notes`  (Month = first day of month date)
### 12. Lessons_Learned (header row 5): `Date | Category | Lesson Learned | Improvement Action | Owner | Status | Notes` (Status `Open`,`Actioned`)
### 13. Bank_Holidays (header row 5): `Date | Name`  — England & Wales 2026 + 2027 (copy from `ATS - KPIs - Data Collation.xlsx` › `Bank Holidays`: 2026-01-01, 04-03, 04-06, 05-04, 05-25, 08-31, 12-25, 12-28; 2027-01-01, 03-26, 03-29, 05-03, 05-31, 08-30, 12-27, 12-28) and note "extend yearly".
### 14. Lists — one list per column with header in row 1: `Priority`,`Severity`,`Test Type`,`Defect Status`,`Found In`,`TSR Status`,`RAG`,`Yes/No`,`Phase`,`Resource Status`,`Lesson Status`,`Topic`,`KPI`.
Topic list (canonical): `Legacy Penalties – Cycle 1 (Individuals)`, `Legacy Penalties – Cycle 2`, `Sign-Up & Reg – Cycle 1`, `Sign-Up & Reg – Cycle 2`, `Manage All Unpaid Agents Capacitors`, `Duplicate/Mixed Records & Lockdown Cases`, `Manage Designatory Details & Signals`, `Manage Designatory Details & Signals – No Repayment Signal`, `SA Digital Data – Enrolment/Registration Support`, `Notice to File (NTF)`, `Business Income Sources`, `Test Automation`, `Programme (cross-cutting)`.
Leave spare blank rows in each list so users can add values (validation ranges should cover e.g. rows 2–40).

## KPI RULES the dashboard implements (for the Guide's RAG table)
| KPI | Rule (source) |
|---|---|
| Risk Based Coverage | per risk level coverage = Covered ÷ Requirements; Green ≥ target; Amber ≥ target − tolerance; else Red. Overall = worst of High & Medium. |
| Production Defect Leakage (quarterly) | defects with Found In = Production in the quarter-to-date. Red if P1 ≥ 1 or P2 ≥ 5 or P3 ≥ 6; Amber if P2 in 1–4 or P3 in 1–5; else Green. "Not in Production" when Go-Live blank. |
| Defect Detection Effectiveness | DDE = Testing-found ÷ (Testing-found + Production-found). Also raised/rejected by priority; rejection % Green ≤5%, Amber ≤10%, else Red. |
| Defect Aging | working-day aging buckets; SLA compliance = share resolved within SLA days; Green ≥ 80%, Amber ≥ 60%, else Red (assumption). |
| Automation Coverage (quarterly) | Green if ≥ target or improved vs previous period, else Amber. |
| Milestone Delivery | per-milestone rule above; overall Red if any Red; Amber if share Amber > 20%; else Green. |
| Environment Availability | downtime hours: monthly Green ≤24, Amber ≤48 else Red; weekly thresholds in Config. |
| TSR Impact Assessment | Red if any TSR working days > SLA (10) else Green. |
| Customer Satisfaction (quarterly) | Green if score ≥ pass-mark and not declining; Red if < pass-mark and not improving; else Amber. "Not Requested" if blank. |
| Resource Position | Actual ÷ Planned: Green ≥ 100%, Amber ≥ 90%, else Red (assumption). |
| Demand Forecast | Planned FTE vs Available FTE for next 3 months: Green if available ≥ planned; Amber gap ≤ 1 FTE; else Red (assumption). |
| Commercial Performance | Forecast vs Budget: Green ≤ budget; Amber ≤ budget +5%; else Red (assumption). |
| RAID Health | Red if any open risk/issue rated Very High (score ≥ 20); Amber if any High (15–19); else Green (assumption). |
