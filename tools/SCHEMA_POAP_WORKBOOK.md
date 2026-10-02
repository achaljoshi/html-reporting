# ATS POAP (Plan on a Page) workbook — schema + JS module contract

File names: `ATS_POAP_Template.xlsx` (blank + example rows) and `ATS_POAP_SAMPLE.xlsx` (seeded from the user's real `ATS - POAP.xlsx`).

## Conventions (same as the KPI workbook)
* Row 1 = title (bold 14pt navy `1B2A4A`), row 2 = purpose, row 3 = HOW TO FILL, row 4 = blank (or group band), **row 5 = header, data from row 6**.
* Parser finds columns by header text (case-insensitive) – keep header text exact. Blank template has ONE grey-italic example row whose `Notes` = `EXAMPLE – delete this row`; **parser skips rows whose Notes start with "EXAMPLE"**.
* Inputs: yellow `FFF6CC` + blue font; calc cells: grey `F2F5F9`. Header navy `1B2A4A` white bold. Arial. Dates real Excel dates `yyyy-mm-dd`. Dropdown validation from `Lists`. Freeze panes + autofilter. Pre-format ~300 rows in the blank template. Calc formulas run through LibreOffice recalc (script `/Users/storezadeveloper/Library/Application Support/Claude/local-agent-mode-sessions/skills-plugin/883e640c-50a9-4746-af99-dadaac066d55/02026406-e98e-46d3-a908-2e2adc27f33f/skills/xlsx/scripts/recalc.py <file> 60`) → 0 errors.

## Sheets (exact names, this order)
1. `Guide`  2. `POAP_Config`  3. `POAP_Plan`  4. `POAP_Milestones`  5. `P2P_Matrix`  6. `Lists`

### POAP_Config  (header row 4: `Setting | Value | Description`; key col A, value col B; then two tables below)
Keys: `Program Name`, `Plan Start Date`, `Plan End Date`, `Status Date (blank = today)`, `Data Mode` (`SAMPLE`/`LIVE`).
Then a block whose first cell (col A) = `PILLARS` followed by header row `Pillar | Colour (hex, no #)` and rows (e.g. `Pillar 1`,`Pillar 2`,`Pillar 3`,`Penalties P2P (WBS)`, `Cross-cutting`).
Then a block whose first cell = `ACTIVITY TYPES` followed by header `Activity Type | Colour (hex, no #)` and rows. Canonical activity types (use these names): `Test Prep`, `Dependencies`, `Regression`, `SIT`, `Smoke / P2P`, `Defect Retest`, `TCR / Sign-off`, `E2E`, `NFR`, `Automation`, `Holiday / Freeze`, `Other`.

### POAP_Plan — one row per bar on the roadmap (header row 5)
```
A ID (unique, e.g. P1-001) | B Pillar | C Topic (swimlane) | D Activity Type | E Item (bar label) |
F Start | G End | H Baseline Start | I Baseline End | J % Complete (0-1) | K Status | L RAG | M Owner |
N Depends On (comma-separated IDs) | O Scope / Notes | P Duration (Days) calc =IF(AND(F<>"",G<>""),G-F+1,"") |
Q Slip (Days) calc =IF(AND(G<>"",I<>""),G-I,"") | R Notes
```
Status list: `Not Started`,`In Progress`,`Complete`,`Delayed`,`On Hold`,`Dates TBC`. RAG list `Green`,`Amber`,`Red`.
(Column O = scope text; R = Notes used for the EXAMPLE flag.)

### POAP_Milestones (header row 5)
`A ID | B Pillar | C Topic | D Milestone | E Date | F Type | G Status | H Notes`   Type list: `SIT Start`,`SIT End`,`Sign-off / TCR`,`Environment`,`Dependency`,`Go-Live`,`Gate`,`Other`; Status: `Planned`,`Achieved`,`At Risk`,`Missed`.

### P2P_Matrix — point-to-point integration test matrix (header row 5)
`A Flow ID | B Topic | C Sub Process | D System Path (use " → " between systems) | E Test No | F Parent Test | G Document | H Planned Date | I Status | J 2xx (success) | K 3xx (redirect) | L 4xx (client error) | M 5xx (server error) | N Notes`
Status list: `Not Started`,`In Progress`,`Complete`,`Blocked`,`Dates TBC`. Each of J–M list: `Not Run`,`Pass`,`Fail`,`Blocked`,`N/A`.

### Lists: one list per column, header in row 1: `Pillar`,`Activity Type`,`Status`,`RAG`,`Milestone Type`,`Milestone Status`,`P2P Status`,`P2P Result`. Leave spare rows.

## Parsed JSON (what `ATS.poap.parse(wb)` returns; JSON-serialisable, dates as `YYYY-MM-DD` strings)
```js
{ config:{ program, start, end, statusDate, mode, pillars:[{name,color}], types:[{name,color}] },
  plan:[{ id,pillar,topic,type,item,start,end,baseStart,baseEnd,pct,status,rag,owner,dependsOn:[ids],notes }],
  milestones:[{ id,pillar,topic,name,date,type,status,notes }],
  p2p:[{ id,topic,subProcess,path,pathNodes:[..],testNo,parent,doc,date,status,r2xx,r3xx,r4xx,r5xx,notes }] }
```

## JS module contract  (file `assets/js/poap.js` + `assets/css/poap.css`; plain non-module script; must run from `file://`; NO network, NO CDN)
```js
window.ATS = window.ATS || {};
ATS.poap = {
  detect(wb)  // SheetJS workbook -> true if it has a sheet named POAP_Plan
  parse(wb)   // SheetJS workbook -> JSON above (or null)
  mount(rootEl)           // called once, builds static DOM inside rootEl
  render(data, ctx)       // (re)render; data = parsed JSON; ctx = { today: Date, sampleMode: boolean }
}
```
Globals available: `XLSX` (SheetJS 0.18.5) and `Chart` (Chart.js 4.4.1, optional). Reuse existing base CSS classes/variables from `assets/css/style.css` (`--navy --teal --amber --red --green --line --bg --sub --ink`, `.card .card-pad .card-head .card-title .card-sub .chip(.green/.amber/.red/.grey) .btn .btn-primary .btn-outline .btn-sm .filter-chip(.active) .toggle-pill .search-box .grid .grid-2/.grid-3/.grid-4 .table-wrap table.data-table`). Do NOT edit style.css; put all new CSS in `poap.css`, every selector prefixed `.poap-`.
