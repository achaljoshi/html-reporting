# ATS Program — Executive Dashboard

A fully offline, interactive dashboard for presenting the ATS program's monthly
portfolio data (KPIs + TSR/SLA performance) to clients. It runs entirely on
your machine — **no server, no internet connection, no upload of any kind.**
Everything is read live from Excel files you keep in a local folder.

---

## 1. What's in this folder

```
ATS_Dashboard/
├── index.html                        ← double-click this to open the dashboard
├── assets/                           ← app code + bundled libraries (offline, no CDN)
├── templates/
│   └── ATS_Portfolio_Template.xlsx   ← blank template — circulate this to portfolio leads
└── ATS_Data/                         ← put completed monthly Excel files here
    ├── 2026-08/                      ← one sample month, pre-filled with demo data
    │   ├── Self Assessment.xlsx
    │   ├── PAYE.xlsx
    │   └── ... (16 portfolios)
    └── 2026-09/                      ← a second sample month, so you can see
        └── ...                         month-over-month comparison working out of the box
```

The `ATS_Data` folder already contains two months of **realistic sample data**
for all 16 portfolios so you can try the dashboard immediately, before your
team has filled in a single real file. Once real data starts coming in,
either overwrite these sample files or delete the `ATS_Data` folder's
contents and start fresh — the dashboard doesn't care, it just reads
whatever `.xlsx` files it finds.

## 2. Opening the dashboard

Just double-click **`index.html`**. It opens in your default browser
(Chrome or Edge recommended — Chromium-based browsers support the folder
picker used to load data). No install, no server, works on a flight with
Wi-Fi off.

The first time, click **Load Data Folder** and select the **`ATS_Data`**
folder itself (not a subfolder inside it). The dashboard scans it for any
subfolder named like `2026-09` and any `.xlsx` file inside — it automatically
figures out which months and portfolios you have.

Because browsers require a click to grant folder access, you'll need to
click **Load / Refresh Data** again any time you add a new month's files —
the dashboard remembers the last thing you loaded (via your browser's local
storage) between sessions, so day-to-day you only need to reload when new
data arrives.

## 3. The monthly cycle

1. At the start of each reporting cycle, send each portfolio lead a fresh
   copy of **`templates/ATS_Portfolio_Template.xlsx`**.
2. They fill in the **yellow cells only** on both tabs:
   - **KPI_Summary** — the 8 quality KPIs (Defect Leakage, Automated Testing
     %, Customer Satisfaction, Risk-Based Test Coverage, Defect Detection,
     Test Execution Downtime, Defect Turnaround Time, Critical Test Milestone
     Delays). Target and Actual are entered; RAG status is calculated
     automatically.
   - **TSR** — one row per Test Scope Request received that month. They only
     enter the TSR ID, description, dates received/responded, resource count,
     and (only if the TSR is still open) a Pending Sub-Status
     (Approved / Query Raised / Impacted / Impacting). **SLA Due Date,
     Response Status (SLA Met / Overdue / Pending) and the Breach flag are
     all calculated automatically from the dates** — nobody has to manually
     judge whether something breached, which is exactly the accuracy
     guarantee you need for a fined-if-wrong contractual metric.
3. They save the file as `<Portfolio Name>.xlsx` (e.g. `Self Assessment.xlsx`)
   and send it back to you (or drop it in a shared location you control).
4. You place each file into `ATS_Data/<YYYY-MM>/`, creating that month's
   folder the first time (e.g. `ATS_Data/2026-10/`).
5. Open the dashboard, click **Load / Refresh Data**, select the `ATS_Data`
   folder again. The new month appears in the month selector, and every
   chart, tile and table can now show **"vs Prior Month"** deltas
   automatically.

You never need to touch the HTML/JS files for this — the whole monthly cycle
is "collect Excel files → drop them in a dated folder → reload."

## 4. Using it in front of a client

### The four views (left sidebar)

- **Overview** — program-wide KPI tiles, an SLA outcome donut you can click
  to drill into the TSR register, and a "portfolios needing attention" list
  sorted by breach count. Click any of the 16 portfolio cards to drill in.
- **Portfolios** — pick any portfolio from the dropdown to see its KPI
  scorecard (8 gauges), SLA donut, pending-TSR breakdown chart, a trend line
  across every month you've loaded, and the full TSR log (searchable,
  sortable, filterable). **Click any chart segment or bar to filter the
  table beneath it** — e.g. click the red "SLA Overdue" slice to instantly
  see just the breaches.
- **Trends & Comparison** — three things:
  1. A program-level trend line (SLA compliance & KPI health across every
     loaded month).
  2. A portfolio leaderboard ranked by SLA compliance for the selected month.
  3. **KPI Deep-Dive** — instead of one generic chart with a metric dropdown,
     each of the 8 KPIs gets its own panel, visualized the way that metric
     is actually used:
     | Metric | Visual | Why |
     |---|---|---|
     | Defect Leakage | Ranked bar (best→worst) | Simple ranking |
     | Level of Automated Testing | Bullet chart (pale target + bold actual) | Target-vs-actual is the point |
     | Customer Satisfaction | Medal leaderboard (🥇🥈🥉) | Reads like a score, not a metric |
     | Risk-Based Test Coverage | Stacked bar (Covered vs Gap to 100%) | Coverage is "how much is left" |
     | Defect Detection Rate | Dot-strip ruler, per-portfolio target tick | Shows spread across the program |
     | Test Execution Downtime | Ranked bar (fewest→most hours) | Simple ranking, time-themed |
     | Defect Turnaround Time | Ruler with diamond marker | Same day-scale idea, distinct shape |
     | Critical Test Milestone Delays | Heatmap tile grid | Values are tiny counts (0-3) — a bar chart would waste space |

     Every bar, tile, dot and leaderboard row is clickable — it jumps straight
     to that portfolio's detail view.
- **TSR Register** — every TSR across every portfolio in one filterable,
  sortable table, useful for a "here's every breach this month, across the
  whole program" view. Each row's portfolio name links back to that
  portfolio's detail view.

### Other controls (top bar)

- **Month selector** — switch which reporting month is on screen.
- **vs Prior Month** toggle — turns on ▲/▼ delta badges everywhere (KPI
  tiles, gauges) comparing the selected month to the one before it.
- **Present** — hides the sidebar/controls and enlarges key numbers, useful
  when screen-sharing or presenting on a projector. **Exit Present** returns
  to normal.
- **Print / PDF** — uses your browser's print dialog with a print-friendly
  stylesheet, if you want to hand out a PDF snapshot after the meeting.

### A note on filters

Filtering the TSR table (by status, search, or a chart click) carries across
from the Portfolio view into the TSR Register if you navigate there next —
this is intentional, so a chart-driven drill-down ("show me this portfolio's
breaches" → "show me *every* portfolio's breaches") stays consistent. The
active filter is always shown as a chip you can click to change or clear.

## 5. Customizing

- **Portfolio names**: just rename the `.xlsx` files — the dashboard reads
  the portfolio name from the filename (and cross-checks it against the
  "Portfolio Name" cell inside the workbook).
- **KPI targets**: each portfolio's target values live inside its own
  workbook (`KPI_Summary` sheet, column F) — set them per portfolio as
  negotiated with the client.
- **SLA window**: the template assumes a 5-calendar-day response SLA (per
  your contract). If that ever changes, open
  `templates/ATS_Portfolio_Template.xlsx`, edit the `SLA Due Date` formula
  in the `TSR` sheet (currently `=Date Received + 5`), and re-circulate the
  updated template.
- **Colors/branding**: `assets/css/style.css` — the palette is defined as
  CSS variables at the top of the file if you want to match a specific
  brand.

## 6. Data privacy

Nothing in this dashboard makes a network request. The two libraries it
uses (SheetJS for reading Excel files, Chart.js for charts) are bundled
locally in `assets/lib/` — open `index.html` with Wi-Fi off and it will
work identically. All parsed data is cached only in your own browser's
local storage on this machine; nothing is uploaded anywhere.
