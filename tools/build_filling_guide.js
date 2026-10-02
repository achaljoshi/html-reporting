#!/usr/bin/env node
// Builds docs/HOW_TO_FILL_DATA.md and docs/ATS_Data_Filling_Guide.docx from ONE content definition.
// usage: NODE_PATH=$(npm root -g) node tools/build_filling_guide.js
const fs = require('fs'), path = require('path');
const D = require('docx');
const root = path.resolve(__dirname, '..');
fs.mkdirSync(path.join(root, 'docs'), { recursive: true });

// ---------------------------------------------------------------- content blocks
const B = [];
const h1 = (t) => B.push({ k: 'h1', t }), h2 = (t) => B.push({ k: 'h2', t }), h3 = (t) => B.push({ k: 'h3', t });
const p = (t) => B.push({ k: 'p', t }), ul = (a) => B.push({ k: 'ul', a }), ol = (a) => B.push({ k: 'ol', a });
const note = (t, kind) => B.push({ k: 'note', t, kind: kind || 'info' }), code = (t) => B.push({ k: 'code', t });
const table = (head, rows, w) => B.push({ k: 'table', head, rows, w });
const pb = () => B.push({ k: 'pb' });

B.push({ k: 'title', t: 'ATS Dashboard — How to Fill the Data', sub: 'Step-by-step guide for portfolio teams and the programme lead' });

h1('1. The idea in 60 seconds');
ul([
  '**One workbook per portfolio.** Every portfolio fills the *same* Excel workbook (`ATS_Weekly_Data.xlsx`, copied from `templates/ATS_Weekly_Data_Template.xlsx`).',
  '**Fill it weekly, never overwrite history.** Each Friday you add a new dated row (and update any defects, TSRs or milestones that changed). Old rows stay — that is what builds the trends.',
  '**Nothing is entered twice.** The Weekly Report, the Monthly Council pack, the Programme Overview and the PowerPoint export are all calculated from the same rows. A month is simply the weeks that end in it.',
  '**Optional extras per portfolio:** a POAP workbook (the plan on a page) and a RAID log (your existing RAID file, unchanged).',
]);
note('Look at a finished example first: open `samples/ATS_Program_SAMPLE/Self Assessment/ATS_Weekly_Data.xlsx` (and its `ATS_POAP.xlsx` and `ATS_RAID_Log.xlsx`) next to this guide. All sample data is fictional.');

h1('2. What you get in the box');
table(['File / folder', 'What it is', 'Who uses it'], [
  ['`templates/ATS_Weekly_Data_Template.xlsx`', 'Blank weekly workbook with an example row on every sheet and a built-in **Guide** sheet.', 'Every portfolio team'],
  ['`templates/ATS_POAP_Template.xlsx`', 'Blank plan-on-a-page workbook (plan bars, milestones, integration flows).', 'Portfolios that want a roadmap'],
  ['`templates/ATS_RAID_Log_Template.xlsx`', 'Blank RAID log with **Risks, Issues, Assumptions, Dependencies and Decisions** sheets, dropdowns and a user guide.', 'Portfolios that keep a RAID log'],
  ['`samples/ATS_Program_SAMPLE/`', 'Five fictional portfolios (Self Assessment, PAYE, VAT, Customs Declaration Service, Child Benefit). Self Assessment also has a POAP and a RAID log.', 'Everyone — to see how a filled workbook looks'],
  ['`ATS_Program/`', 'Your real data folder. Put one subfolder per portfolio here.', 'Programme lead'],
  ['`index.html`', 'The dashboard. Double-click to open (Chrome or Edge).', 'Everyone'],
], [3200, 4438, 2000]);

h1('3. One-time set-up for each portfolio');
ol([
  'Create a folder for the portfolio inside `ATS_Program/`, e.g. `ATS_Program/Self Assessment/`.',
  'Copy `templates/ATS_Weekly_Data_Template.xlsx` into it and rename it `ATS_Weekly_Data.xlsx`.',
  'Open the **Config** sheet. Set **Portfolio Name** (must be unique across portfolios — it is the name in the dashboard dropdown), **Programme**, **Program Name**, **Project / Phase**, **TSR Reference**, the managers, **Total Budget (£)** and the programme dates. Set **Data Mode** to `LIVE`.',
  'Review the thresholds further down Config (SLA days, downtime limits, coverage targets). The defaults are sensible; change them only if your contract says otherwise.',
  'On the **Lists** sheet, replace the example **Topic** values with your portfolio\'s own topics / releases / workstreams (they feed the dropdowns).',
  '**Delete the grey EXAMPLE rows** on each data sheet (the row whose Notes column says "EXAMPLE – delete this row"). Rows marked EXAMPLE are ignored by the dashboard, but delete them anyway.',
  'Optional: copy `templates/ATS_POAP_Template.xlsx` as `ATS_POAP.xlsx` and `templates/ATS_RAID_Log_Template.xlsx` as `ATS_RAID_Log.xlsx` into the same folder (or drop in your existing RAID file).',
]);
code('ATS_Program/\n├── Self Assessment/\n│   ├── ATS_Weekly_Data.xlsx      <- required\n│   ├── ATS_POAP.xlsx             <- optional\n│   └── ATS_RAID_Log.xlsx         <- optional (your existing RAID format)\n├── PAYE/\n│   └── ATS_Weekly_Data.xlsx\n└── VAT/\n    └── ATS_Weekly_Data.xlsx');
note('Prefer one flat folder? Name the files `ATS_Weekly_Data - <Portfolio Name>.xlsx`. A POAP or RAID file is matched to its portfolio when the portfolio name appears in its file name.');

h1('4. The weekly routine (every Friday, about 5 minutes)');
table(['Sheet', 'What to do', 'How many rows'], [
  ['**Weekly_Snapshot**', 'Add ONE new row for this week (Week Ending = the Friday). Enter downtime, coverage, automation, resources, spend and your overall RAG with a one-paragraph summary.', '1 per week'],
  ['**Topic_Progress**', 'Add one row per active topic with cumulative test-case counts to date.', '1 per topic per week'],
  ['**Defect_Log**', 'Add new defects; update status / Resolved Date of existing ones. Never delete a defect.', 'as it happens'],
  ['**TSR_Log**', 'Add new TSRs when received; fill Returned Date and Hold Days when you respond.', 'as it happens'],
  ['**Milestones**', 'Fill the Completed Date when a milestone is done; adjust the Forecast Date if it moves.', 'as it happens'],
  ['**Readiness**', 'Add one row per topic with six RAG areas and the Road to Green.', '1 per topic per week (optional)'],
  ['**Commentary**', 'Optional: your own wording for any KPI. If you leave it empty the dashboard drafts the text for you.', 'optional'],
], [1900, 6338, 1400]);
p('**Monthly / quarterly items** (when they change): CSAT score in `Weekly_Snapshot`, `Demand_Forecast`, `Resourcing`, `Lessons_Learned`.');
p('Then save the workbook. The programme lead opens the dashboard and clicks **Refresh data** (or **Load program folder** the first time).');

note('**Golden rules:** (1) week ending dates must be **Fridays**; (2) yellow cells are yours, grey cells are formulas — never type over grey; (3) counts in Weekly_Snapshot and Topic_Progress are **cumulative to date**, not "this week only"; (4) never delete history; (5) keep every portfolio\'s **Portfolio Name** unique.', 'warn');

h1('5. Sheet-by-sheet reference');
p('Yellow = you type. Grey = calculated. The header text must stay exactly as in the template — the dashboard finds columns by their header.');

h2('5.1 Config — settings (once)');
table(['Setting', 'What to enter', 'Example'], [
  ['Portfolio Name', 'The name shown in the dashboard. Unique per portfolio.', 'Self Assessment'],
  ['Programme', 'Groups portfolios under one heading.', 'HMRC ATS'],
  ['Program Name', 'Full title used in page headers.', 'HMRC ATS – Self Assessment'],
  ['Project / Phase, TSR Reference', 'Shown on commercial pages.', 'Phase 1 / TSR_PR021877-REQ031204'],
  ['Programme Test Manager, Test Manager', 'Names for the resourcing page.', 'role or name'],
  ['Total Budget (£)', 'Approved budget; the commercial KPI compares the forecast with it.', '760000'],
  ['Production Go-Live Date', 'Leave blank until live. While blank, "Production Defect Leakage" shows *Not in Production*.', '(blank)'],
  ['Data Mode', '`LIVE` for real data, `SAMPLE` for demo data (adds a SAMPLE tag in the dashboard).', 'LIVE'],
  ['Thresholds', 'TSR SLA (10 working days), defect SLA (5), rejection %, downtime limits, coverage targets, milestone window, automation target, CSAT pass mark, execution blocked limits.', 'defaults provided'],
], [2600, 4838, 2200]);

h2('5.2 Weekly_Snapshot — one row per week');
table(['Column(s)', 'What to enter', 'Example (Self Assessment, W/E 2 Oct 2026)'], [
  ['Week Ending', 'The Friday of the reporting week.', '2026-10-02'],
  ['Planned Test Days', 'Working days testing was planned that week (bank-holiday weeks: 4).', '5'],
  ['Downtime Hours', 'Hours test execution was suspended (environment outage etc.). **Availability %** is calculated.', '0'],
  ['High / Medium / Low Risk – Requirements', 'Total requirements tagged at that risk level (cumulative).', '72 / 100 / 48'],
  ['… – Test Cases', 'Test cases written against that risk level.', '126 / 189 / 105'],
  ['… – Covered Reqs', 'Requirements fully covered by scripts. **Coverage %** is calculated (covered ÷ requirements).', '72 / 97 / 30'],
  ['Automatable / Automated Test Data Setups', 'How many test-data set-ups can be automated, and how many are.', '150 / 108'],
  ['Automatable / Automated Test Cases', 'Same for test-case execution automation.', '200 / 150'],
  ['CSAT Score (1-5), CSAT Responses', 'Only in a week a score was collected; otherwise leave blank ("Not requested").', '(blank)'],
  ['Resources Planned / Actual (FTE)', 'Planned vs actual full-time equivalents in place.', '6 / 6'],
  ['Cumulative Spend (£), Forecast at Completion (£)', 'Spend to date and the latest forecast.', '608000 / 760000'],
  ['Overall RAG', 'Your judgement: Green, Amber or Red. Shown alongside the RAG the dashboard calculates.', 'Green'],
  ['Overall Executive Summary', 'One paragraph for the council / report.', '"Portfolio remains Green: execution on plan…"'],
], [2900, 4338, 2400]);

h2('5.3 Topic_Progress — one row per topic per week (cumulative)');
table(['Column', 'What to enter', 'Example'], [
  ['Week Ending, Topic', 'The Friday and the topic / release / workstream (pick from the dropdown).', '2026-10-02, Self Assessment – Digital Filing'],
  ['Phase', 'Test Prep, Smoke, SIT, Regression, E2E, NFR or TCR / Sign-off.', 'SIT'],
  ['TCs Planned / Designed', 'Test cases planned in scope, and designed (scripted) so far.', '70 / 63'],
  ['Passed / Failed / Blocked', 'Cumulative results to date. **Executed**, **% Executed** and **% Designed** are calculated.', '4 / 0 / 0'],
  ['Comments', 'Short note.', '4 executed (4 passed, 0 failed); 0 blocked.'],
], [2300, 4938, 2400]);

h2('5.4 Defect_Log — one row per defect');
table(['Column', 'What to enter', 'Example'], [
  ['Defect ID, Summary', 'From your defect tool (Jira etc.).', 'SA-101, Tax calculation – message queued but not delivered'],
  ['Topic, Test Type', 'Topic dropdown; Smoke, SIT, E2E, NFR, Regression or Production.', 'Self Assessment – Release 1, Smoke'],
  ['Priority', 'P1 High, P2 Medium or P3 Low.', 'P2 Medium'],
  ['Severity', '1 Urgent … 6 Trivial (see the definitions in the Guide).', '3 - High'],
  ['Raised Date', 'Date raised.', '2026-08-05'],
  ['Resolved Date', 'The date work **stops** (fix and confirmation test done, or the date it was rejected). Blank while open.', '2026-08-06'],
  ['Status', 'Open, In Progress, Re-Open, Resolved, Closed, Rejected or Deferred.', 'Closed'],
  ['Found In', 'Testing or Production. Production defects drive the leakage KPI.', 'Testing'],
  ['Owner / DG, TCs Blocked, Triage Comments', 'Who is fixing it, test cases it blocks, latest note.', 'Data Team, 0, Fix delivered and confirmed.'],
  ['Aging, Aging Bucket (grey)', 'Calculated in **working days**, bank holidays excluded: 0-2, 3-5, 6-10, >10 days.', '2, 0-2 Days'],
], [2300, 4938, 2400]);
note('A **Closed** or **Resolved** defect must have a Resolved Date — otherwise its aging cannot be measured (the Data Quality page will flag it).');

h2('5.5 TSR_Log — one row per Test Scope Request');
table(['Column', 'What to enter', 'Example'], [
  ['TSR Ref, Description, Topic', 'Reference, short description, topic.', 'PR022310-REQ032588'],
  ['Received Date', 'Date the TSR reached you.', '2026-09-28'],
  ['Returned Date', 'Date you sent it back to the front desk with commercials complete. Blank while open (it then ages to today).', '(blank)'],
  ['Hold Days (pending clarification)', 'Working days you were waiting for clarification — these are subtracted.', '0'],
  ['Status', 'Open, Returned or Withdrawn (withdrawn TSRs are ignored).', 'Open'],
  ['Work Days, Over SLA? (grey)', 'Calculated: working days minus hold days; "Yes" if above the limit in Config (default 10).', '5, No'],
], [2900, 4338, 2400]);

h2('5.6 Milestones — one row per milestone');
table(['Column', 'What to enter', 'Example'], [
  ['Milestone, Topic, Owner', 'Name, topic (or Programme), owner role.', 'Test Completion Report, Programme (cross-cutting)'],
  ['Due Date', 'The agreed (baseline) date.', '2026-12-18'],
  ['Forecast Date', 'Optional: the current forecast if it has moved.', '(blank)'],
  ['Completed Date', 'Fill when done. Completed after the Due Date = late.', '(blank)'],
  ['Critical?', 'Yes for milestones that count towards "critical milestone delay".', 'No'],
  ['Delay, Status, RAG (grey)', 'Calculated: Completed / Completed Late / Overdue / On Schedule; Amber in the last 5 days before the due date, Red once overdue.', '0, On Schedule, Green'],
], [2300, 4938, 2400]);

h2('5.7 Readiness — one row per topic per week (optional)');
p('Columns: **Week Ending, Topic, Planned SIT Start, Overall RAG**, then for each of six areas — *Requirements & Arch Intent, Blockers, Resources, Environments, Test Scripts & Test Data, Knowledge Transfer* — a **RAG** and a **Note**, and finally **Road to Green** (what has to happen to turn the topic Green). Example: Overall Green, every area Green, note "Requirements baselined and reviewed."');

h2('5.8 Commentary — your own words (optional)');
p('One row per KPI per week: **Week Ending, KPI** (pick from the list), **Key Highlights, Area of Concern, Actions Underway, Executive Commentary**, optional **RAG Override** and **Current Position Override**. Put each bullet on its own line (Alt+Enter). Anything you type here replaces the auto-drafted text for that KPI and is labelled *From Excel* in the dashboard; blank fields keep the auto draft.');

h2('5.9 Small tables');
table(['Sheet', 'Columns', 'Notes'], [
  ['Resourcing', 'Role, Name, Organisation, Allocation %, Start Date, End Date, Status', 'Status: Active, Planned or Left. Use role names if you prefer not to list people.'],
  ['Demand_Forecast', 'Month (first day), Planned FTE, Available FTE, Notes', 'A gap of more than 1 FTE in the next 3 months turns the Demand KPI Red.'],
  ['Lessons_Learned', 'Date, Category, Lesson Learned, Improvement Action, Owner, Status', 'Shown on the "Lessons Learned" council page.'],
  ['Bank_Holidays', 'Date, Name', 'Used for all working-day maths. Add next year\'s dates each January.'],
  ['Lists', 'Dropdown values', 'Edit Topic here; leave the rest unless you add a new value.'],
], [2000, 4438, 3200]);

h1('6. POAP — plan on a page');
p('Fill `ATS_POAP.xlsx` once and update dates and status as the plan moves. The dashboard turns it into a roadmap (Gantt), a one-page poster and an integration-flow explorer.');
table(['Sheet', 'One row per', 'Key columns'], [
  ['POAP_Config', 'setting', 'Program Name, Plan Start/End Date, **Status Date** (blank = today), Data Mode, Pillars and Activity Types with colours'],
  ['POAP_Plan', 'bar on the roadmap', 'ID, Pillar, Topic (swim-lane), Activity Type, Item, **Start, End**, **Baseline Start, Baseline End**, % Complete, Status, RAG, Owner, Depends On, Scope / Notes'],
  ['POAP_Milestones', 'milestone', 'ID, Pillar, Topic, Milestone, Date, Type (SIT Start, SIT End, Sign-off / TCR, Go-Live, Gate…), Status (Planned, Achieved, At Risk, Missed)'],
  ['P2P_Matrix', 'integration flow', 'Flow ID, Topic, Sub Process, **System Path** (`System A → System B → System C`), Test No, Planned Date, Status, results for 2xx / 3xx / 4xx / 5xx'],
], [1900, 2000, 5738]);
h3('How to fill POAP_Plan');
ul([
  '**Start / End** are the *current* plan. **Baseline Start / End** are the *originally agreed* dates — fill them once and never change them. The dashboard shows the gap as a ghost bar and a slip badge such as **+14d** ("Slip (Days)" is calculated).',
  '**Status:** Not Started, In Progress, Complete, Delayed, On Hold or Dates TBC. For *Dates TBC* leave Start and End blank.',
  '**% Complete** is 0–100% (type 0.5 or 50%).',
  '**Depends On:** the IDs this bar waits for, comma-separated (e.g. `A-003, A-004`). The roadmap draws the dependency arrows.',
  '**Activity Type** decides the bar colour. Use the dropdown: Test Prep, Dependencies, Regression, SIT, Smoke / P2P, Defect Retest, TCR / Sign-off, E2E, NFR, Automation, Holiday / Freeze, Other.',
  'Rows whose Notes start with `EXAMPLE` are ignored — delete the example row once you have real rows.',
]);

h1('7. RAID log');
p('Start a new log from `templates/ATS_RAID_Log_Template.xlsx` (or keep using your existing RAID file — the dashboard finds columns by their header names). Put it in the portfolio\'s folder as `ATS_RAID_Log.xlsx`. It has five log sheets, each with the header on row 10 and data from row 11.');
table(['Sheet', 'One row per', 'Key columns'], [
  ['Risks', 'risk', 'Summary Title, Likelihood, Impact (score = L × I → Very Low … Very High), Status (Open / Mitigated / Closed), Mitigation, Owner, Review Date'],
  ['Issues', 'issue', 'Summary Title, Priority, Severity, Status (Open / Resolution in progress / Resolved / Closed), Target and Actual Resolution Date'],
  ['Assumptions', 'assumption', 'Description, Confidence Level, Validation Action, Validation Due Date, Status (Unconfirmed / Confirmed Correct / Confirmed Incorrect)'],
  ['Dependencies', 'dependency', 'Description, Dependency For / From, Type, Date Required, Priority, Status (Open / Closed)'],
  ['**Decisions**', 'decision', 'Summary Title, Decision / Description, Category, Decision Maker / Forum, **Date Decided**, Rationale, Impact of Decision, **Status (Pending / Approved / Rejected / Superseded)**, Owner, Review Date, Linked Risk / Issue ID'],
], [1800, 1400, 6438]);
p('The dashboard gives each sheet its own page under **Risks, Issues & Lessons** (RAID Overview, Risks, Issues, Assumptions, Dependencies, Decisions, Lessons Learned). Every page has its own counts, a clickable matrix (risk likelihood × impact, issue priority × severity, assumption confidence × status, dependency priority × when needed, decision category × status), charts, auto-drafted highlights and the full list of open items with closed ones tucked away. Decisions marked *Pending* count as open and appear under "Actions Underway"; decisions whose *Date Decided* falls in the month or week are counted in the highlights. An old RAID file without a Decisions sheet is fine — that part simply shows "No decisions logged". **RAID Health** is Red when a Very High item is open and Amber when a High item is open. A closed risk or a resolved issue drops out automatically; set *Archived = Yes* to hide any row.');

h1('8. After you fill: load and check');
ol([
  'Open `index.html`. Click **Load program folder** and choose `ATS_Program` (first time), or **Refresh data** after later edits.',
  'Use the **Portfolio** dropdown (top bar) to switch between portfolios, or **Programme Overview** to see them all.',
  'Open **Appendix → Data Quality** in the Monthly Council or Weekly Report. It lists problems such as a closed defect without a Resolved Date, a TSR returned before it was received, a week ending that is not a Friday, text typed into a number cell, a POAP bar that ends before it starts or depends on an ID that does not exist, and RAID items with duplicate IDs or no rating.',
  'Use **Export PowerPoint** for slides (built on the client PowerPoint theme; charts and tables stay editable), **Export PDF** for a document you can read or print (the Monthly Council and Weekly Report export every page of the pack), or **Copy as picture** on any single page.',
]);

h1('9. How the numbers are worked out');
table(['KPI', 'Rule (thresholds are editable in Config)'], [
  ['Risk Based Coverage', 'Covered ÷ requirements per risk level. Green at or above target; Amber within 2 points; else Red. Overall = worst of High and Medium.'],
  ['Production Defect Leakage (quarterly)', 'Red if any P1, or 5+ P2, or 6+ P3 found in production in the quarter; Amber for any P2/P3; else Green.'],
  ['Defect Detection Effectiveness', 'Defects found in testing ÷ (testing + production). Rejection rate Green ≤ 5%, Amber ≤ 10%.'],
  ['Defect Aging', 'Working-day buckets; Green if at least 80% resolved within 5 working days, Amber at least 60%, else Red.'],
  ['Automation Coverage (quarterly)', 'Green if at or above target, or improving on the previous period.'],
  ['Milestone Delivery', 'Red if any milestone is overdue or late; Amber if more than 20% are Amber.'],
  ['Environment Availability', '1 − downtime ÷ (test days × 8). Monthly: Green ≤ 24h, Amber ≤ 48h. Weekly limits are in Config.'],
  ['TSR Impact Assessment', 'Red if any TSR takes more than 10 working days (hold days excluded). Open TSRs age to today.'],
  ['Customer Satisfaction (quarterly)', 'Pass mark 3 of 5; Green if at or above it and not declining.'],
  ['Resource / Demand / Commercial', 'Actual ÷ planned FTE; FTE gap in the next 3 months; forecast vs budget (Amber within 5% over).'],
  ['RAID Health', 'Red with a Very High item open; Amber with a High item open.'],
  ['Test Execution Progress', 'Blocked ÷ planned: Green ≤ 5%, Amber ≤ 20%, else Red.'],
], [3000, 6638]);
h3('How weekly rows become a month');
ul([
  '**Added up:** planned test days and downtime hours.',
  '**Latest week in the period:** coverage, automation, test execution, resources, spend, forecast.',
  '**Filtered by date:** defects, TSRs, milestones.',
  'A week belongs to the month it **ends** in. Every month is calculated as of its last day, so last month\'s pack never changes when you add new rows.',
]);

h1('10. Common mistakes');
table(['Mistake', 'What happens', 'Fix'], [
  ['Week Ending is not a Friday', 'Flagged on the Data Quality page; the week may land in the wrong month.', 'Use the Friday of that week.'],
  ['Typed "this week only" numbers in Weekly_Snapshot / Topic_Progress', 'Coverage and execution fall instead of grow.', 'Enter cumulative totals to date.'],
  ['Deleted old defect/TSR rows', 'Past months change and trends break.', 'Keep history; change Status instead.'],
  ['Closed a defect without a Resolved Date', 'Aging cannot be measured.', 'Add the date work stopped.'],
  ['Two workbooks with the same Portfolio Name', 'One portfolio disappears from the dropdown.', 'Make the names unique.'],
  ['Typed over a grey formula cell', 'Calculated column stops updating.', 'Restore the formula from the template.'],
  ['Left the EXAMPLE row in', 'Ignored by the dashboard, but confusing.', 'Delete it.'],
  ['Forgot to click Refresh data', 'The dashboard still shows the previous load.', 'Click Refresh data.'],
], [3300, 3338, 3000]);

h1('11. Glossary');
table(['Term', 'Meaning'], [
  ['RAG', 'Red / Amber / Green status.'], ['TSR', 'Test Scope Request — a request to estimate or scope testing, answered within an SLA.'],
  ['SLA', 'The agreed response time (default 10 working days for a TSR, 5 for defect turnaround).'], ['SIT / E2E / NFR', 'System Integration Test, End-to-End test, Non-Functional test.'],
  ['P2P', 'Point-to-point integration test between two systems.'], ['POAP', 'Plan on a Page — the programme roadmap.'],
  ['RAID', 'Risks, Assumptions, Issues, Dependencies.'], ['FTE', 'Full-time equivalent.'],
], [2000, 7638]);

// ---------------------------------------------------------------- markdown output
const md = [];
const mdInline = (t) => t;
B.forEach((b) => {
  if (b.k === 'title') md.push(`# ${b.t}\n\n*${b.sub}*\n`);
  else if (b.k === 'h1') md.push(`\n## ${b.t}\n`);
  else if (b.k === 'h2') md.push(`\n### ${b.t}\n`);
  else if (b.k === 'h3') md.push(`\n#### ${b.t}\n`);
  else if (b.k === 'p') md.push(mdInline(b.t) + '\n');
  else if (b.k === 'ul') md.push(b.a.map((x) => '- ' + x).join('\n') + '\n');
  else if (b.k === 'ol') md.push(b.a.map((x, i) => `${i + 1}. ${x}`).join('\n') + '\n');
  else if (b.k === 'note') md.push(`> ${b.kind === 'warn' ? '⚠️ ' : 'ℹ️ '}${b.t}\n`);
  else if (b.k === 'code') md.push('```\n' + b.t + '\n```\n');
  else if (b.k === 'table') md.push(`| ${b.head.join(' | ')} |\n|${b.head.map(() => '---').join('|')}|\n` + b.rows.map((r) => `| ${r.map((c) => c.replace(/\|/g, '\\|')).join(' | ')} |`).join('\n') + '\n');
});
fs.writeFileSync(path.join(root, 'docs', 'HOW_TO_FILL_DATA.md'), md.join('\n'));

// ---------------------------------------------------------------- docx output
const NAVY = '1B2A4A', TEAL = '0A7C78', FONT = 'Arial';
function runs(text, base) {
  base = base || {};
  const out = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`|\*[^*]+\*)/g;
  text.split(re).filter((s) => s !== '').forEach((s) => {
    if (s.startsWith('**')) out.push(new D.TextRun(Object.assign({ text: s.slice(2, -2), bold: true, font: FONT }, base)));
    else if (s.startsWith('`')) out.push(new D.TextRun(Object.assign({ text: s.slice(1, -1), font: 'Consolas', size: (base.size || 21) - 2, color: '9A3412' }, { })));
    else if (s.startsWith('*') && s.length > 2) out.push(new D.TextRun(Object.assign({ text: s.slice(1, -1), italics: true, font: FONT }, base)));
    else out.push(new D.TextRun(Object.assign({ text: s, font: FONT }, base)));
  });
  return out;
}
const W = 9638;
const border = { style: D.BorderStyle.SINGLE, size: 4, color: 'D0D5DD' };
const borders = { top: border, bottom: border, left: border, right: border };
const children = [];
const olRefs = [];
B.forEach((b) => {
  if (b.k === 'title') {
    children.push(new D.Paragraph({ spacing: { before: 600, after: 120 }, children: [new D.TextRun({ text: b.t, bold: true, size: 48, color: NAVY, font: FONT })] }));
    children.push(new D.Paragraph({ spacing: { after: 300 }, border: { bottom: { style: D.BorderStyle.SINGLE, size: 8, color: TEAL, space: 8 } }, children: [new D.TextRun({ text: b.sub, size: 24, color: '667085', font: FONT })] }));
    children.push(new D.Paragraph({ spacing: { after: 120 }, children: [new D.TextRun({ text: 'Contents', bold: true, size: 26, color: NAVY, font: FONT })] }));
    B.filter((x) => x.k === 'h1').forEach((x) => children.push(new D.Paragraph({ spacing: { after: 60 }, indent: { left: 200 }, children: [new D.TextRun({ text: x.t, size: 21, color: '344054', font: FONT })] })));
    children.push(new D.Paragraph({ children: [new D.PageBreak()] }));
  } else if (b.k === 'h1') children.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_1, spacing: { before: 360, after: 140 }, keepNext: true, children: [new D.TextRun({ text: b.t, bold: true, size: 32, color: NAVY, font: FONT })] }));
  else if (b.k === 'h2') children.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_2, spacing: { before: 260, after: 100 }, keepNext: true, children: [new D.TextRun({ text: b.t, bold: true, size: 26, color: TEAL, font: FONT })] }));
  else if (b.k === 'h3') children.push(new D.Paragraph({ heading: D.HeadingLevel.HEADING_3, spacing: { before: 200, after: 80 }, keepNext: true, children: [new D.TextRun({ text: b.t, bold: true, size: 22, color: '344054', font: FONT })] }));
  else if (b.k === 'p') children.push(new D.Paragraph({ spacing: { after: 120, line: 300 }, children: runs(b.t, { size: 21 }) }));
  else if (b.k === 'ul') b.a.forEach((x) => children.push(new D.Paragraph({ numbering: { reference: 'bul', level: 0 }, spacing: { after: 70, line: 290 }, children: runs(x, { size: 21 }) })));
  else if (b.k === 'ol') { const ref = 'num' + children.length; olRefs.push(ref); b.a.forEach((x) => children.push(new D.Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { after: 70, line: 290 }, children: runs(x, { size: 21 }) }))); }
  else if (b.k === 'note') {
    const warn = b.kind === 'warn';
    children.push(new D.Paragraph({ spacing: { before: 100, after: 160, line: 290 }, shading: { type: D.ShadingType.CLEAR, fill: warn ? 'FFF7E6' : 'EAF6F5' }, border: { left: { style: D.BorderStyle.SINGLE, size: 24, color: warn ? 'F5A524' : '0EA5A0', space: 8 } }, indent: { left: 160, right: 100 }, children: runs(b.t, { size: 20 }) }));
  } else if (b.k === 'code') {
    b.t.split('\n').forEach((line, i, a) => children.push(new D.Paragraph({ spacing: { after: 0, before: i === 0 ? 100 : 0 }, shading: { type: D.ShadingType.CLEAR, fill: 'F2F4F7' }, indent: { left: 120, right: 120 }, children: [new D.TextRun({ text: line || ' ', font: 'Consolas', size: 18, color: '1F2937' })] })));
    children.push(new D.Paragraph({ spacing: { after: 120 }, children: [] }));
  } else if (b.k === 'table') {
    const tot = b.w.reduce((a, c) => a + c, 0), ws = b.w.map((x) => Math.round((x / tot) * W));
    ws[ws.length - 1] += W - ws.reduce((a, c) => a + c, 0);
    const cell = (t, i, header, zebra) => new D.TableCell({ width: { size: ws[i], type: D.WidthType.DXA }, borders, margins: { top: 70, bottom: 70, left: 100, right: 100 }, shading: { type: D.ShadingType.CLEAR, fill: header ? NAVY : zebra ? 'F8FAFC' : 'FFFFFF' }, children: [new D.Paragraph({ spacing: { line: 260 }, children: runs(t, header ? { size: 18, bold: true, color: 'FFFFFF' } : { size: 18 }) })] });
    children.push(new D.Table({ width: { size: W, type: D.WidthType.DXA }, columnWidths: ws, rows: [new D.TableRow({ tableHeader: true, children: b.head.map((h, i) => cell(h, i, true)) })].concat(b.rows.map((r, ri) => new D.TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, i, false, ri % 2 === 1)) }))) }));
    children.push(new D.Paragraph({ spacing: { after: 140 }, children: [] }));
  }
});
const doc = new D.Document({
  creator: 'ATS Dashboard', title: 'ATS Dashboard — How to Fill the Data',
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: 'bul', levels: [{ level: 0, format: D.LevelFormat.BULLET, text: '•', alignment: D.AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }]
    .concat(olRefs.map((r) => ({ reference: r, levels: [{ level: 0, format: D.LevelFormat.DECIMAL, text: '%1.', alignment: D.AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 320 } } } }] }))) },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
    footers: { default: new D.Footer({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: 'ATS Dashboard — How to Fill the Data  ·  page ', size: 16, color: '98A2B3', font: FONT }), new D.TextRun({ children: [D.PageNumber.CURRENT], size: 16, color: '98A2B3', font: FONT })] })] }) },
    children,
  }],
});
D.Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(path.join(root, 'docs', 'ATS_Data_Filling_Guide.docx'), buf); console.log('wrote docs/HOW_TO_FILL_DATA.md and docs/ATS_Data_Filling_Guide.docx'); });
