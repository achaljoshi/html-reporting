/* ============================================================
   ATS Program — Executive Dashboard
   Fully offline. Reads Excel files selected from a local folder.
   No network calls, no server, no external dependencies at runtime.
   ============================================================ */

(function () {
  "use strict";

  if (window.Chart && window.Chart.registerables) {
    Chart.register(...Chart.registerables);
  }

  const RAG_COLOR = { Green: "#22c55e", Amber: "#f5a524", Red: "#ef4444", "": "#c9d1e0" };
  const RAG_SCORE = { Green: 100, Amber: 60, Red: 20 };
  const STATUS_COLOR = { "SLA Met": "#22c55e", "SLA Overdue": "#ef4444", "Pending": "#f5a524" };
  const PENDING_COLOR = { Approved: "#0ea5a0", "Query Raised": "#6366f1", Impacted: "#f5a524", Impacting: "#ef4444" };
  const LS_KEY = "ats_dashboard_dataset_v1";
  const MONTH_RE = /^(\d{4})-(\d{2})$/;

  // ---------------------------------------------------------------
  // State
  // ---------------------------------------------------------------
  const state = {
    months: [],              // sorted ascending ["2026-08","2026-09"]
    dataByMonth: {},          // { "2026-09": { "Self Assessment": {kpi:[], tsr:[], meta:{}} } }
    selectedMonth: null,
    compareOn: false,
    currentView: "overview",
    selectedPortfolio: null,
    tsrFilter: { status: "All", search: "", pendingSub: null, scope: "all" }, // scope: all | portfolio
    tsrSort: { key: "received", dir: "desc" },
    sourceLabel: "",
    charts: {},               // named Chart.js instances, destroyed/recreated on redraw
  };

  // ---------------------------------------------------------------
  // Utils
  // ---------------------------------------------------------------
  function $(sel, root) { return (root || document).querySelector(sel); }
  function $all(sel, root) { return Array.from((root || document).querySelectorAll(sel)); }
  function el(tag, attrs, children) {
    const e = document.createElement(tag);
    if (attrs) for (const k in attrs) {
      if (k === "class") e.className = attrs[k];
      else if (k === "html") e.innerHTML = attrs[k];
      else if (k.startsWith("on") && typeof attrs[k] === "function") e.addEventListener(k.slice(2), attrs[k]);
      else e.setAttribute(k, attrs[k]);
    }
    (children || []).forEach(c => { if (c != null) e.appendChild(typeof c === "string" ? document.createTextNode(c) : c); });
    return e;
  }
  function fmtMonth(m) {
    if (!m) return "—";
    const match = MONTH_RE.exec(m);
    if (!match) return m;
    const d = new Date(Number(match[1]), Number(match[2]) - 1, 1);
    return d.toLocaleDateString(undefined, { month: "long", year: "numeric" });
  }
  function fmtMonthShort(m) {
    const match = MONTH_RE.exec(m);
    if (!match) return m;
    const d = new Date(Number(match[1]), Number(match[2]) - 1, 1);
    return d.toLocaleDateString(undefined, { month: "short", year: "2-digit" });
  }
  function fmtDate(d) {
    if (!d) return "—";
    const dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(dt)) return "—";
    return dt.toLocaleDateString(undefined, { day: "2-digit", month: "short", year: "numeric" });
  }
  function pct(n, d) { return d ? Math.round((n / d) * 1000) / 10 : 0; }
  function round1(n) { return Math.round(n * 10) / 10; }
  function ragChip(rag, extraText) {
    const cls = { Green: "green", Amber: "amber", Red: "red" }[rag] || "grey";
    return `<span class="chip ${cls}">${extraText || rag || "—"}</span>`;
  }
  function prevMonth(m) {
    const idx = state.months.indexOf(m);
    return idx > 0 ? state.months[idx - 1] : null;
  }
  function deltaBadge(curr, prev, higherIsBetter, unit) {
    if (prev == null || curr == null || isNaN(prev)) return `<span class="kpi-delta flat">No prior data</span>`;
    const diff = round1(curr - prev);
    if (diff === 0) return `<span class="kpi-delta flat">No change vs last month</span>`;
    const good = higherIsBetter ? diff > 0 : diff < 0;
    const arrow = diff > 0 ? "▲" : "▼";
    const cls = good ? "up" : "down";
    return `<span class="kpi-delta ${cls}">${arrow} ${Math.abs(diff)}${unit || ""} vs ${fmtMonthShort(prevMonth(state.selectedMonth) || "")}</span>`;
  }

  // ---------------------------------------------------------------
  // Excel parsing
  // ---------------------------------------------------------------
  function cellVal(ws, addr) {
    const c = ws[addr];
    return c ? c.v : undefined;
  }

  function parseKpiSheet(ws) {
    if (!ws) return null;
    const portfolio = cellVal(ws, "C4");
    const month = cellVal(ws, "F4");
    const preparedBy = cellVal(ws, "C5");
    const metrics = [];
    for (let r = 8; r <= 15; r++) {
      const name = cellVal(ws, "C" + r);
      if (!name) continue;
      metrics.push({
        name: String(name),
        unit: cellVal(ws, "D" + r) || "",
        direction: cellVal(ws, "E" + r) || "Higher",
        target: Number(cellVal(ws, "F" + r)) || 0,
        actual: cellVal(ws, "G" + r) === undefined || cellVal(ws, "G" + r) === "" ? null : Number(cellVal(ws, "G" + r)),
        rag: cellVal(ws, "H" + r) || "",
        comments: cellVal(ws, "I" + r) || "",
      });
    }
    return { portfolio: portfolio ? String(portfolio) : null, month: month ? String(month) : null, preparedBy, metrics };
  }

  function excelDateToJS(v) {
    if (v == null || v === "") return null;
    if (v instanceof Date) return v;
    if (typeof v === "number") {
      // SheetJS with cellDates:true normally returns Date objects already;
      // fallback conversion just in case.
      const utc = XLSX.SSF ? null : null;
      const epoch = new Date(Date.UTC(1899, 11, 30));
      return new Date(epoch.getTime() + v * 86400000);
    }
    const d = new Date(v);
    return isNaN(d) ? null : d;
  }

  function parseTsrSheet(ws, portfolioName, month) {
    if (!ws) return [];
    const ref = ws["!ref"];
    let maxRow = 200;
    if (ref) { const range = XLSX.utils.decode_range(ref); maxRow = range.e.r + 1; }
    const rows = [];
    for (let r = 6; r <= maxRow; r++) {
      const tsrId = cellVal(ws, "C" + r);
      const received = excelDateToJS(cellVal(ws, "E" + r));
      if (!tsrId && !received) continue;
      const responded = excelDateToJS(cellVal(ws, "G" + r));
      const due = excelDateToJS(cellVal(ws, "F" + r));
      let status = cellVal(ws, "I" + r);
      if (!status) status = responded ? (due && responded <= due ? "SLA Met" : "SLA Overdue") : "Pending";
      rows.push({
        id: tsrId ? String(tsrId) : ("TSR-" + r),
        desc: cellVal(ws, "D" + r) || "",
        received, due, responded,
        resources: Number(cellVal(ws, "H" + r)) || 0,
        status: String(status),
        pendingSub: cellVal(ws, "J" + r) || "",
        breach: (cellVal(ws, "K" + r) || (status === "SLA Overdue" ? "Yes" : "No")),
        comments: cellVal(ws, "L" + r) || "",
        portfolio: portfolioName,
        month: month,
      });
    }
    return rows;
  }

  function readWorkbookFile(file, relPath) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = function (e) {
        try {
          const data = new Uint8Array(e.target.result);
          const wb = XLSX.read(data, { type: "array", cellDates: true, cellNF: false });
          resolve({ wb, relPath, file });
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  async function ingestFiles(fileList) {
    const files = Array.from(fileList).filter(f => /\.xlsx$/i.test(f.name) && !f.name.startsWith("~$"));
    if (!files.length) {
      alert("No .xlsx files found in the selected folder. Pick the top-level folder that contains your YYYY-MM month subfolders (e.g. ATS_Data).");
      return;
    }
    const dataByMonth = {};
    const monthsSet = new Set();
    let rootName = "";
    let okCount = 0, skipCount = 0;

    for (const file of files) {
      const relPath = file.webkitRelativePath || file.name;
      const parts = relPath.split("/");
      // Expect .../<root>/<YYYY-MM>/<Portfolio Name>.xlsx  (root folder is optional depth)
      let monthFolder = null, portfolioFile = null;
      for (let i = 0; i < parts.length - 1; i++) {
        if (MONTH_RE.test(parts[i])) { monthFolder = parts[i]; portfolioFile = parts[parts.length - 1]; break; }
      }
      if (!rootName && parts.length > 1) rootName = parts[0];
      if (!monthFolder) { skipCount++; continue; }

      try {
        const { wb } = await readWorkbookFile(file);
        const kpiSheetName = wb.SheetNames.find(n => /kpi/i.test(n)) || wb.SheetNames[1] || wb.SheetNames[0];
        const tsrSheetName = wb.SheetNames.find(n => /tsr/i.test(n)) || wb.SheetNames[2] || wb.SheetNames[1];
        const kpi = parseKpiSheet(wb.Sheets[kpiSheetName]);
        const portfolioName = (kpi && kpi.portfolio) || portfolioFile.replace(/\.xlsx$/i, "");
        const tsr = parseTsrSheet(wb.Sheets[tsrSheetName], portfolioName, monthFolder);

        monthsSet.add(monthFolder);
        dataByMonth[monthFolder] = dataByMonth[monthFolder] || {};
        dataByMonth[monthFolder][portfolioName] = {
          meta: { portfolio: portfolioName, month: monthFolder, preparedBy: kpi ? kpi.preparedBy : "", fileName: file.name },
          kpi: kpi ? kpi.metrics : [],
          tsr,
        };
        okCount++;
      } catch (err) {
        console.error("Failed to parse", relPath, err);
        skipCount++;
      }
    }

    if (!okCount) {
      alert("Could not read any valid portfolio workbooks. Make sure the folder structure is:\n<root>/<YYYY-MM>/<Portfolio Name>.xlsx\nand each file has KPI_Summary and TSR sheets.");
      return;
    }

    state.months = Array.from(monthsSet).sort();
    state.dataByMonth = dataByMonth;
    state.selectedMonth = state.months[state.months.length - 1];
    state.sourceLabel = rootName || "Selected folder";
    state.selectedPortfolio = null;

    persist();
    alertBanner(`Loaded ${okCount} portfolio file(s) across ${state.months.length} month(s).` + (skipCount ? ` (${skipCount} file(s) skipped)` : ""));
    populateMonthSelect();
    renderAll();
  }

  function alertBanner(msg) {
    const status = $("#legacy-data-status");
    status.innerHTML = `<b>${msg}</b>`;
  }

  function persist() {
    try {
      const payload = { months: state.months, dataByMonth: state.dataByMonth, sourceLabel: state.sourceLabel, savedAt: new Date().toISOString() };
      localStorage.setItem(LS_KEY, JSON.stringify(payload, (k, v) => v));
    } catch (e) { console.warn("Could not persist dataset to localStorage", e); }
  }

  function loadPersisted() {
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (!raw) return false;
      const payload = JSON.parse(raw);
      // revive date strings back into Date objects
      Object.values(payload.dataByMonth).forEach(portfolios => {
        Object.values(portfolios).forEach(p => {
          p.tsr.forEach(row => {
            row.received = row.received ? new Date(row.received) : null;
            row.due = row.due ? new Date(row.due) : null;
            row.responded = row.responded ? new Date(row.responded) : null;
          });
        });
      });
      state.months = payload.months;
      state.dataByMonth = payload.dataByMonth;
      state.sourceLabel = payload.sourceLabel;
      state.selectedMonth = state.months[state.months.length - 1];
      updateDataStatus(payload.savedAt);
      return true;
    } catch (e) { console.warn("Could not load persisted dataset", e); return false; }
  }

  function updateDataStatus(savedAt) {
    const portfolios = state.selectedMonth ? Object.keys(state.dataByMonth[state.selectedMonth] || {}) : [];
    const when = savedAt ? new Date(savedAt).toLocaleString() : "";
    $("#legacy-data-status").innerHTML = `<b>${portfolios.length}</b> portfolios · <b>${state.months.length}</b> month(s) loaded${when ? `<br>Last loaded: ${when}` : ""}`;
  }

  // ---------------------------------------------------------------
  // Aggregation
  // ---------------------------------------------------------------
  function portfolioNames(month) {
    return Object.keys(state.dataByMonth[month] || {}).sort();
  }

  function getPortfolio(month, name) {
    return (state.dataByMonth[month] || {})[name];
  }

  function tsrSummary(tsrRows) {
    const total = tsrRows.length;
    const met = tsrRows.filter(r => r.status === "SLA Met").length;
    const overdue = tsrRows.filter(r => r.status === "SLA Overdue").length;
    const pending = tsrRows.filter(r => r.status === "Pending").length;
    const pendingBreakdown = {};
    tsrRows.filter(r => r.status === "Pending" && r.pendingSub).forEach(r => {
      pendingBreakdown[r.pendingSub] = (pendingBreakdown[r.pendingSub] || 0) + 1;
    });
    const resources = tsrRows.reduce((s, r) => s + (r.resources || 0), 0);
    return { total, met, overdue, pending, pendingBreakdown, slaPct: pct(met, total), breachPct: pct(overdue, total), resources };
  }

  function kpiHealthScore(kpiRows) {
    const scored = kpiRows.filter(m => m.rag && RAG_SCORE[m.rag] != null);
    if (!scored.length) return null;
    const sum = scored.reduce((s, m) => s + RAG_SCORE[m.rag], 0);
    return Math.round(sum / scored.length);
  }

  function programSummary(month) {
    const names = portfolioNames(month);
    let allTsr = [];
    let healthScores = [];
    names.forEach(n => {
      const p = getPortfolio(month, n);
      allTsr = allTsr.concat(p.tsr);
      const hs = kpiHealthScore(p.kpi);
      if (hs != null) healthScores.push(hs);
    });
    const summary = tsrSummary(allTsr);
    const health = healthScores.length ? Math.round(healthScores.reduce((a, b) => a + b, 0) / healthScores.length) : null;
    return { ...summary, health, portfolioCount: names.length, allTsr };
  }

  function allTsrForMonth(month) {
    const names = portfolioNames(month);
    let rows = [];
    names.forEach(n => rows = rows.concat(getPortfolio(month, n).tsr));
    return rows;
  }

  // ---------------------------------------------------------------
  // Chart helpers
  // ---------------------------------------------------------------
  function destroyChart(key) {
    if (state.charts[key]) { state.charts[key].destroy(); delete state.charts[key]; }
  }
  function makeChart(key, ctx, config) {
    destroyChart(key);
    state.charts[key] = new Chart(ctx, config);
    return state.charts[key];
  }

  // ---------------------------------------------------------------
  // View: Overview
  // ---------------------------------------------------------------
  function renderOverview() {
    const root = $("#view-overview");
    if (!state.selectedMonth) { root.innerHTML = emptyStateHTML(); wireEmptyState(); return; }

    const month = state.selectedMonth;
    const prev = prevMonth(month);
    const summary = programSummary(month);
    const prevSummary = prev ? programSummary(prev) : null;
    const names = portfolioNames(month);

    root.innerHTML = `
      <div class="grid grid-4" style="margin-bottom:18px;">
        ${kpiTile("Total TSRs Received", summary.total, prevSummary ? prevSummary.total : null, true, "", "#0ea5a0")}
        ${kpiTile("SLA Compliance", summary.slaPct + "%", prevSummary ? prevSummary.slaPct : null, true, "%", "#22c55e", summary.slaPct)}
        ${kpiTile("Contract Breaches (Fines)", summary.overdue, prevSummary ? prevSummary.overdue : null, false, "", "#ef4444")}
        ${kpiTile("Program KPI Health", summary.health != null ? summary.health : "—", prevSummary ? prevSummary.health : null, true, "", "#6366f1")}
      </div>

      <div class="grid grid-3" style="margin-bottom:18px; align-items:stretch;">
        <div class="card card-pad">
          <div class="card-head"><div><div class="card-title">TSR Outcome Split</div><div class="card-sub">${fmtMonth(month)} · click a slice to view in TSR Register</div></div></div>
          <div class="chart-box" style="height:220px;"><canvas id="chart-overview-donut"></canvas></div>
        </div>
        <div class="card card-pad" style="grid-column:span 2;">
          <div class="card-head"><div><div class="card-title">Portfolios Needing Attention</div><div class="card-sub">Ranked by breach count &amp; overdue TSRs this month</div></div></div>
          <div id="risk-list"></div>
        </div>
      </div>

      <div class="card card-pad">
        <div class="card-head">
          <div><div class="card-title">All Portfolios</div><div class="card-sub">${names.length} portfolios · click any card to drill in</div></div>
        </div>
        <div class="grid grid-4" id="portfolio-grid"></div>
      </div>
    `;

    // Donut
    const ctx = $("#chart-overview-donut").getContext("2d");
    makeChart("overviewDonut", ctx, {
      type: "doughnut",
      data: {
        labels: ["SLA Met", "SLA Overdue", "Pending"],
        datasets: [{ data: [summary.met, summary.overdue, summary.pending], backgroundColor: [STATUS_COLOR["SLA Met"], STATUS_COLOR["SLA Overdue"], STATUS_COLOR["Pending"]], borderWidth: 0 }]
      },
      options: {
        maintainAspectRatio: false,
        cutout: "68%",
        plugins: { legend: { position: "bottom", labels: { boxWidth: 10, font: { size: 11 } } } },
        onClick: (evt, elements) => {
          if (!elements.length) return;
          const idx = elements[0].index;
          const label = ["SLA Met", "SLA Overdue", "Pending"][idx];
          openTsrRegister({ scope: "all", status: label });
        }
      }
    });

    // Risk list
    const riskRoot = $("#risk-list");
    const riskRows = names.map(n => {
      const p = getPortfolio(month, n);
      const s = tsrSummary(p.tsr);
      return { name: n, overdue: s.overdue, total: s.total, slaPct: s.slaPct };
    }).filter(r => r.overdue > 0).sort((a, b) => b.overdue - a.overdue).slice(0, 6);
    if (!riskRows.length) {
      riskRoot.innerHTML = `<p class="text-sub" style="font-size:13px;">No SLA breaches recorded this month across any portfolio.</p>`;
    } else {
      riskRoot.innerHTML = riskRows.map(r => `
        <div class="leader-row" data-portfolio="${escapeAttr(r.name)}" style="cursor:pointer;">
          <div class="leader-rank" style="background:#fef3f2;color:#b42318;">${r.overdue}</div>
          <div class="leader-name">${r.name}</div>
          <div class="leader-bar"><span style="width:${100 - r.slaPct}%;background:#ef4444;"></span></div>
          <div class="leader-val">${r.slaPct}% met</div>
        </div>`).join("");
      $all("[data-portfolio]", riskRoot).forEach(row => row.addEventListener("click", () => openPortfolio(row.getAttribute("data-portfolio"))));
    }

    // Portfolio grid
    const grid = $("#portfolio-grid");
    grid.innerHTML = names.map(n => portfolioCardHTML(month, n)).join("");
    $all(".portfolio-card", grid).forEach(card => card.addEventListener("click", () => openPortfolio(card.getAttribute("data-portfolio"))));
  }

  function kpiTile(label, value, prevVal, higherIsBetter, unit, color, barPct) {
    const delta = typeof prevVal === "number" && typeof value !== "string"
      ? deltaBadge(Number(value), prevVal, higherIsBetter, unit)
      : (typeof value === "string" ? "" : deltaBadge(value, prevVal, higherIsBetter, unit));
    return `
      <div class="kpi-tile">
        <div class="accent" style="background:${color};"></div>
        <div class="kpi-label">${label}</div>
        <div class="kpi-value">${value}</div>
        ${state.compareOn ? delta : `<div class="kpi-foot">&nbsp;</div>`}
      </div>`;
  }

  function portfolioCardHTML(month, name) {
    const p = getPortfolio(month, name);
    const s = tsrSummary(p.tsr);
    const health = kpiHealthScore(p.kpi);
    const healthColor = health == null ? "#c9d1e0" : health >= 80 ? "#22c55e" : health >= 55 ? "#f5a524" : "#ef4444";
    const metPct = s.total ? (s.met / s.total) * 100 : 0;
    const overduePct = s.total ? (s.overdue / s.total) * 100 : 0;
    const pendPct = s.total ? (s.pending / s.total) * 100 : 0;
    return `
      <div class="portfolio-card" data-portfolio="${escapeAttr(name)}">
        <div class="pc-top">
          <div class="pc-name">${name}</div>
          <span class="chip ${health == null ? "grey" : health >= 80 ? "green" : health >= 55 ? "amber" : "red"}">${health != null ? health + " health" : "no KPI"}</span>
        </div>
        <div class="pc-metrics">
          <div class="pc-metric"><div class="num">${s.total}</div><div class="lbl">TSRs</div></div>
          <div class="pc-metric"><div class="num" style="color:${s.overdue ? '#ef4444' : '#22c55e'}">${s.overdue}</div><div class="lbl">Breaches</div></div>
          <div class="pc-metric"><div class="num">${s.slaPct}%</div><div class="lbl">SLA Met</div></div>
        </div>
        <div class="pc-bar">
          <span style="width:${metPct}%;background:${STATUS_COLOR["SLA Met"]}"></span>
          <span style="width:${overduePct}%;background:${STATUS_COLOR["SLA Overdue"]}"></span>
          <span style="width:${pendPct}%;background:${STATUS_COLOR["Pending"]}"></span>
        </div>
        <div class="pc-foot"><span>View details</span><span>→</span></div>
      </div>`;
  }

  function escapeAttr(s) { return String(s).replace(/"/g, "&quot;"); }
  function escapeHtml(s) { return String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])); }

  // ---------------------------------------------------------------
  // View: Portfolio detail
  // ---------------------------------------------------------------
  function openPortfolio(name) {
    state.selectedPortfolio = name;
    switchView("portfolio");
  }

  function renderPortfolioView() {
    const root = $("#view-portfolio");
    if (!state.selectedMonth) { root.innerHTML = emptyStateHTML(); wireEmptyState(); return; }
    const names = portfolioNames(state.selectedMonth);
    if (!state.selectedPortfolio || !names.includes(state.selectedPortfolio)) state.selectedPortfolio = names[0];
    const name = state.selectedPortfolio;
    if (!name) { root.innerHTML = `<p class="text-sub">No portfolios found for this month.</p>`; return; }

    const month = state.selectedMonth;
    const prev = prevMonth(month);
    const p = getPortfolio(month, name);
    const prevP = prev ? getPortfolio(prev, name) : null;
    const s = tsrSummary(p.tsr);
    const prevS = prevP ? tsrSummary(prevP.tsr) : null;

    root.innerHTML = `
      <div class="toolbar-row" style="justify-content:space-between;">
        <select id="portfolio-switch" style="min-width:220px;font-weight:600;"></select>
        <div style="display:flex;gap:16px;">
          <div class="chip ${s.overdue ? 'red' : 'green'}">${s.overdue} breach${s.overdue === 1 ? "" : "es"} this month</div>
          <div class="chip grey">${s.total} TSRs</div>
        </div>
      </div>

      <div class="grid grid-4" style="margin-bottom:18px;">
        ${kpiTile("TSRs Received", s.total, prevS ? prevS.total : null, true)}
        ${kpiTile("SLA Compliance", s.slaPct + "%", prevS ? prevS.slaPct : null, true, "%")}
        ${kpiTile("Overdue (Breach)", s.overdue, prevS ? prevS.overdue : null, false)}
        ${kpiTile("Pending", s.pending, prevS ? prevS.pending : null, false)}
      </div>

      <div class="card card-pad" style="margin-bottom:18px;">
        <div class="card-head"><div><div class="card-title">KPI Scorecard</div><div class="card-sub">Target vs actual · ${fmtMonth(month)}</div></div></div>
        <div class="grid grid-4" id="gauge-grid"></div>
      </div>

      <div class="grid grid-3" style="margin-bottom:18px;align-items:stretch;">
        <div class="card card-pad">
          <div class="card-head"><div class="card-title">TSR Outcome</div></div>
          <div class="chart-box" style="height:200px;"><canvas id="chart-portfolio-donut"></canvas></div>
        </div>
        <div class="card card-pad">
          <div class="card-head"><div class="card-title">Pending Breakdown</div></div>
          <div class="chart-box" style="height:200px;"><canvas id="chart-portfolio-pending"></canvas></div>
        </div>
        <div class="card card-pad">
          <div class="card-head"><div class="card-title">SLA Compliance Trend</div><div class="card-sub">across loaded months</div></div>
          <div class="chart-box" style="height:200px;"><canvas id="chart-portfolio-trend"></canvas></div>
        </div>
      </div>

      <div class="card card-pad">
        <div class="card-head">
          <div><div class="card-title">TSR Log — ${name}</div><div class="card-sub">${fmtMonth(month)}</div></div>
        </div>
        <div id="portfolio-tsr-table"></div>
      </div>
    `;

    const sel = $("#portfolio-switch");
    sel.innerHTML = names.map(n => `<option value="${escapeAttr(n)}" ${n === name ? "selected" : ""}>${n}</option>`).join("");
    sel.addEventListener("change", () => { state.selectedPortfolio = sel.value; state.tsrFilter.status = "All"; state.tsrFilter.search = ""; renderPortfolioView(); });

    renderGauges(p.kpi, prevP ? prevP.kpi : null);
    renderPortfolioDonut(s);
    renderPendingBreakdown(s);
    renderPortfolioTrend(name);
    renderPortfolioTsrTable(p.tsr, "#portfolio-tsr-table");
  }

  function renderGauges(kpiRows, prevKpiRows) {
    const grid = $("#gauge-grid");
    if (!kpiRows.length) { grid.innerHTML = `<p class="text-sub">No KPI data in this workbook.</p>`; return; }
    grid.innerHTML = kpiRows.map(m => {
      const prevM = prevKpiRows ? prevKpiRows.find(x => x.name === m.name) : null;
      const higherIsBetter = m.direction === "Higher";
      const ratio = m.target ? Math.min(1.3, (m.actual || 0) / m.target) : 0;
      const fillPct = Math.min(100, Math.max(4, (higherIsBetter ? ratio : (2 - ratio)) * 50));
      const color = RAG_COLOR[m.rag] || "#c9d1e0";
      let deltaHtml = "";
      if (state.compareOn && prevM && prevM.actual != null && m.actual != null) {
        const diff = round1(m.actual - prevM.actual);
        const good = higherIsBetter ? diff >= 0 : diff <= 0;
        deltaHtml = `<div class="gauge-delta" style="color:${diff === 0 ? '#667085' : (good ? '#067647' : '#b42318')}">${diff === 0 ? "No change" : (diff > 0 ? "▲" : "▼") + " " + Math.abs(diff) + m.unit} vs last month</div>`;
      }
      return `
        <div class="gauge-card">
          <div class="gauge-top">
            <div class="gauge-name">${m.name}</div>
            ${ragChip(m.rag)}
          </div>
          <div class="gauge-values">
            <div class="gauge-actual">${m.actual != null ? m.actual : "—"}${m.unit}</div>
            <div class="gauge-target">target ${m.target}${m.unit} (${m.direction === "Higher" ? "≥" : "≤"})</div>
          </div>
          <div class="gauge-track"><div class="gauge-fill" style="width:${fillPct}%;background:${color};"></div></div>
          ${deltaHtml}
        </div>`;
    }).join("");
  }

  function renderPortfolioDonut(s) {
    const ctx = $("#chart-portfolio-donut").getContext("2d");
    makeChart("portfolioDonut", ctx, {
      type: "doughnut",
      data: { labels: ["SLA Met", "SLA Overdue", "Pending"], datasets: [{ data: [s.met, s.overdue, s.pending], backgroundColor: [STATUS_COLOR["SLA Met"], STATUS_COLOR["SLA Overdue"], STATUS_COLOR["Pending"]], borderWidth: 0 }] },
      options: {
        maintainAspectRatio: false, cutout: "65%",
        plugins: { legend: { position: "bottom", labels: { boxWidth: 10, font: { size: 11 } } } },
        onClick: (evt, elements) => {
          if (!elements.length) return;
          const label = ["SLA Met", "SLA Overdue", "Pending"][elements[0].index];
          state.tsrFilter = { status: label, search: "", scope: "portfolio" };
          renderPortfolioTsrTable(getPortfolio(state.selectedMonth, state.selectedPortfolio).tsr, "#portfolio-tsr-table");
          $("#portfolio-tsr-table").scrollIntoView({ behavior: "smooth", block: "start" });
        }
      }
    });
  }

  function renderPendingBreakdown(s) {
    const labels = Object.keys(PENDING_COLOR);
    const data = labels.map(l => s.pendingBreakdown[l] || 0);
    const ctx = $("#chart-portfolio-pending").getContext("2d");
    if (!data.some(v => v > 0)) {
      ctx.canvas.parentElement.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:var(--sub);font-size:13px;">No pending TSRs</div>`;
      return;
    }
    makeChart("pendingBar", ctx, {
      type: "bar",
      data: { labels, datasets: [{ data, backgroundColor: labels.map(l => PENDING_COLOR[l]), borderRadius: 6, maxBarThickness: 34 }] },
      options: {
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: { y: { beginAtZero: true, ticks: { precision: 0 } }, x: { grid: { display: false } } },
        onClick: (evt, elements) => {
          if (!elements.length) return;
          const label = labels[elements[0].index];
          state.tsrFilter = { status: "Pending", search: "", pendingSub: label, scope: "portfolio" };
          renderPortfolioTsrTable(getPortfolio(state.selectedMonth, state.selectedPortfolio).tsr, "#portfolio-tsr-table");
          $("#portfolio-tsr-table").scrollIntoView({ behavior: "smooth", block: "start" });
        }
      }
    });
  }

  function renderPortfolioTrend(name) {
    const labels = state.months.map(fmtMonthShort);
    const slaData = state.months.map(m => {
      const p = getPortfolio(m, name);
      if (!p) return null;
      return tsrSummary(p.tsr).slaPct;
    });
    const healthData = state.months.map(m => {
      const p = getPortfolio(m, name);
      return p ? kpiHealthScore(p.kpi) : null;
    });
    const ctx = $("#chart-portfolio-trend").getContext("2d");
    makeChart("portfolioTrend", ctx, {
      type: "line",
      data: {
        labels,
        datasets: [
          { label: "SLA Compliance %", data: slaData, borderColor: "#0ea5a0", backgroundColor: "rgba(14,165,160,.12)", fill: true, tension: 0.35, pointRadius: 4 },
          { label: "KPI Health", data: healthData, borderColor: "#6366f1", backgroundColor: "rgba(99,102,241,.08)", fill: true, tension: 0.35, pointRadius: 4 },
        ]
      },
      options: {
        maintainAspectRatio: false,
        plugins: { legend: { position: "bottom", labels: { boxWidth: 10, font: { size: 10.5 } } } },
        scales: { y: { min: 0, max: 100 }, x: { grid: { display: false } } }
      }
    });
  }

  function sortRows(rows, key, dir) {
    const copy = rows.slice();
    copy.sort((a, b) => {
      let av = a[key], bv = b[key];
      if (av instanceof Date || bv instanceof Date) { av = av ? av.getTime() : -Infinity; bv = bv ? bv.getTime() : -Infinity; }
      if (typeof av === "string") av = av.toLowerCase();
      if (typeof bv === "string") bv = bv.toLowerCase();
      if (av == null) av = "";
      if (bv == null) bv = "";
      if (av < bv) return dir === "asc" ? -1 : 1;
      if (av > bv) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return copy;
  }

  function applyTsrFilter(rows, f) {
    let filtered = rows;
    if (f.status && f.status !== "All") filtered = filtered.filter(r => r.status === f.status);
    if (f.pendingSub) filtered = filtered.filter(r => r.pendingSub === f.pendingSub);
    if (f.search) filtered = filtered.filter(r =>
      (r.id + " " + r.desc + " " + r.pendingSub + " " + (r.portfolio || "")).toLowerCase().includes(f.search.toLowerCase())
    );
    return filtered;
  }

  function activeFilterTagHTML(f) {
    if (!f.pendingSub) return "";
    return `<button class="filter-chip active" id="clear-pending-sub" style="background:var(--navy);border-color:var(--navy);color:#fff;">Pending sub-status: ${f.pendingSub} &times;</button>`;
  }

  function renderPortfolioTsrTable(rows, targetSel) {
    const root = $(targetSel);
    const f = state.tsrFilter;
    let filtered = sortRows(applyTsrFilter(rows, f), state.tsrSort.key, state.tsrSort.dir);

    const statuses = ["All", "SLA Met", "SLA Overdue", "Pending"];
    root.innerHTML = `
      <div class="toolbar-row">
        <div class="search-box">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
          <input type="text" id="tsr-search" placeholder="Search TSR ID, description..." value="${escapeAttr(f.search)}">
        </div>
        ${statuses.map(st => `<button class="filter-chip ${f.status === st ? "active" : ""}" data-status="${st}">${st}${st !== "All" ? " (" + rows.filter(r => r.status === st).length + ")" : ""}</button>`).join("")}
        ${activeFilterTagHTML(f)}
      </div>
      <div class="table-wrap">
        <table class="data-table" id="tsr-table">
          <thead><tr>
            ${thSort("id", "TSR ID")}${thSort("desc", "Scope")}${thSort("received", "Received")}${thSort("due", "SLA Due")}${thSort("responded", "Responded")}${thSort("resources", "Resources")}${thSort("status", "Status")}${thSort("pendingSub", "Pending Sub-Status")}
          </tr></thead>
          <tbody>
            ${filtered.map(r => `
              <tr class="${r.status === 'SLA Overdue' ? 'breach-row' : ''}">
                <td>${r.id}</td>
                <td style="white-space:normal;max-width:260px;">${escapeHtml(r.desc)}</td>
                <td>${fmtDate(r.received)}</td>
                <td>${fmtDate(r.due)}</td>
                <td>${r.responded ? fmtDate(r.responded) : "—"}</td>
                <td>${r.resources}</td>
                <td>${statusChip(r.status)}</td>
                <td>${r.pendingSub || "—"}</td>
              </tr>`).join("") || `<tr><td colspan="8" style="text-align:center;color:var(--sub);padding:30px;">No matching TSRs</td></tr>`}
          </tbody>
        </table>
      </div>
    `;

    $("#tsr-search").addEventListener("input", (e) => { state.tsrFilter.search = e.target.value; renderPortfolioTsrTable(rows, targetSel); });
    $all("[data-status]", root).forEach(btn => btn.addEventListener("click", () => {
      state.tsrFilter.status = btn.getAttribute("data-status");
      state.tsrFilter.pendingSub = null;
      renderPortfolioTsrTable(rows, targetSel);
    }));
    const clearBtn = $("#clear-pending-sub", root);
    if (clearBtn) clearBtn.addEventListener("click", () => { state.tsrFilter.pendingSub = null; renderPortfolioTsrTable(rows, targetSel); });
    $all("th[data-key]", root).forEach(th => th.addEventListener("click", () => {
      const key = th.getAttribute("data-key");
      if (state.tsrSort.key === key) state.tsrSort.dir = state.tsrSort.dir === "asc" ? "desc" : "asc";
      else { state.tsrSort.key = key; state.tsrSort.dir = "asc"; }
      renderPortfolioTsrTable(rows, targetSel);
    }));
  }

  function thSort(key, label) {
    const active = state.tsrSort.key === key;
    const arrow = active ? (state.tsrSort.dir === "asc" ? "↑" : "↓") : "";
    return `<th data-key="${key}">${label}${active ? `<span class="arrow">${arrow}</span>` : ""}</th>`;
  }
  function statusChip(status) {
    const cls = status === "SLA Met" ? "green" : status === "SLA Overdue" ? "red" : "amber";
    return `<span class="chip ${cls}">${status}</span>`;
  }

  // ---------------------------------------------------------------
  // View: TSR Register (all portfolios)
  // ---------------------------------------------------------------
  function openTsrRegister(filter) {
    state.tsrFilter = { status: filter.status || "All", search: filter.search || "", pendingSub: filter.pendingSub || null, scope: "all" };
    switchView("tsr");
  }

  function renderTsrView() {
    const root = $("#view-tsr");
    if (!state.selectedMonth) { root.innerHTML = emptyStateHTML(); wireEmptyState(); return; }
    const rows = allTsrForMonth(state.selectedMonth);
    root.innerHTML = `
      <div class="card card-pad">
        <div class="card-head">
          <div><div class="card-title">All Portfolios — TSR Register</div><div class="card-sub">${fmtMonth(state.selectedMonth)} · ${rows.length} total TSRs</div></div>
        </div>
        <div id="tsr-all-table"></div>
      </div>
    `;
    renderAllTsrTable(rows);
  }

  function renderAllTsrTable(rows) {
    const root = $("#tsr-all-table");
    const f = state.tsrFilter;
    let filtered = sortRows(applyTsrFilter(rows, f), state.tsrSort.key, state.tsrSort.dir);
    const statuses = ["All", "SLA Met", "SLA Overdue", "Pending"];

    root.innerHTML = `
      <div class="toolbar-row">
        <div class="search-box">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>
          <input type="text" id="tsr-all-search" placeholder="Search TSR ID, portfolio, description..." value="${escapeAttr(f.search)}">
        </div>
        ${statuses.map(st => `<button class="filter-chip ${f.status === st ? "active" : ""}" data-status="${st}">${st}${st !== "All" ? " (" + rows.filter(r => r.status === st).length + ")" : ""}</button>`).join("")}
        ${activeFilterTagHTML(f)}
      </div>
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr>
            ${thSort("portfolio", "Portfolio")}${thSort("id", "TSR ID")}${thSort("desc", "Scope")}${thSort("received", "Received")}${thSort("due", "SLA Due")}${thSort("responded", "Responded")}${thSort("status", "Status")}${thSort("pendingSub", "Pending Sub-Status")}
          </tr></thead>
          <tbody>
            ${filtered.map(r => `
              <tr class="${r.status === 'SLA Overdue' ? 'breach-row' : ''}">
                <td><a style="cursor:pointer;font-weight:600;color:var(--navy);" data-goto="${escapeAttr(r.portfolio)}">${r.portfolio}</a></td>
                <td>${r.id}</td>
                <td style="white-space:normal;max-width:240px;">${escapeHtml(r.desc)}</td>
                <td>${fmtDate(r.received)}</td>
                <td>${fmtDate(r.due)}</td>
                <td>${r.responded ? fmtDate(r.responded) : "—"}</td>
                <td>${statusChip(r.status)}</td>
                <td>${r.pendingSub || "—"}</td>
              </tr>`).join("") || `<tr><td colspan="8" style="text-align:center;color:var(--sub);padding:30px;">No matching TSRs</td></tr>`}
          </tbody>
        </table>
      </div>
    `;
    $("#tsr-all-search").addEventListener("input", (e) => { state.tsrFilter.search = e.target.value; renderAllTsrTable(rows); });
    $all("[data-status]", root).forEach(btn => btn.addEventListener("click", () => {
      state.tsrFilter.status = btn.getAttribute("data-status");
      state.tsrFilter.pendingSub = null;
      renderAllTsrTable(rows);
    }));
    const clearBtn = $("#clear-pending-sub", root);
    if (clearBtn) clearBtn.addEventListener("click", () => { state.tsrFilter.pendingSub = null; renderAllTsrTable(rows); });
    $all("th[data-key]", root).forEach(th => th.addEventListener("click", () => {
      const key = th.getAttribute("data-key");
      if (state.tsrSort.key === key) state.tsrSort.dir = state.tsrSort.dir === "asc" ? "desc" : "asc";
      else { state.tsrSort.key = key; state.tsrSort.dir = "asc"; }
      renderAllTsrTable(rows);
    }));
    $all("[data-goto]", root).forEach(a => a.addEventListener("click", () => openPortfolio(a.getAttribute("data-goto"))));
  }

  // ---------------------------------------------------------------
  // View: Trends & Comparison
  // ---------------------------------------------------------------
  function renderTrendsView() {
    const root = $("#view-trends");
    if (!state.selectedMonth) { root.innerHTML = emptyStateHTML(); wireEmptyState(); return; }
    const names = portfolioNames(state.selectedMonth);
    const metricNames = [];
    names.forEach(n => getPortfolio(state.selectedMonth, n).kpi.forEach(m => { if (!metricNames.includes(m.name)) metricNames.push(m.name); }));

    root.innerHTML = `
      <div class="grid grid-2" style="margin-bottom:18px; align-items:stretch;">
        <div class="card card-pad">
          <div class="card-head"><div><div class="card-title">Program Trend</div><div class="card-sub">SLA compliance &amp; KPI health across all loaded months</div></div></div>
          <div class="chart-box" style="height:260px;"><canvas id="chart-program-trend"></canvas></div>
        </div>
        <div class="card card-pad">
          <div class="card-head"><div><div class="card-title">Portfolio Leaderboard</div><div class="card-sub">SLA compliance · ${fmtMonth(state.selectedMonth)}</div></div></div>
          <div id="leaderboard" style="max-height:260px;overflow:auto;"></div>
        </div>
      </div>

      <div class="card-title" style="margin:6px 2px 10px;font-size:15.5px;">KPI Deep-Dive — Every Portfolio, Side by Side</div>
      <div class="card-sub" style="margin:-6px 2px 14px;">${fmtMonth(state.selectedMonth)} · each metric is visualized the way it's actually used, not as one generic bar chart</div>
      <div class="grid grid-2" id="kpi-deepdive"></div>
    `;

    // Program trend
    const labels = state.months.map(fmtMonthShort);
    const slaSeries = state.months.map(m => programSummary(m).slaPct);
    const healthSeries = state.months.map(m => programSummary(m).health);
    makeChart("programTrend", $("#chart-program-trend").getContext("2d"), {
      type: "line",
      data: { labels, datasets: [
        { label: "SLA Compliance %", data: slaSeries, borderColor: "#0ea5a0", backgroundColor: "rgba(14,165,160,.12)", fill: true, tension: 0.35, pointRadius: 4 },
        { label: "KPI Health", data: healthSeries, borderColor: "#6366f1", backgroundColor: "rgba(99,102,241,.08)", fill: true, tension: 0.35, pointRadius: 4 },
      ]},
      options: { maintainAspectRatio: false, plugins: { legend: { position: "bottom", labels: { boxWidth: 10, font: { size: 11 } } } }, scales: { y: { min: 0, max: 100 }, x: { grid: { display: false } } } }
    });

    // Leaderboard
    const board = names.map(n => {
      const s = tsrSummary(getPortfolio(state.selectedMonth, n).tsr);
      return { name: n, val: s.slaPct };
    }).sort((a, b) => b.val - a.val);
    $("#leaderboard").innerHTML = board.map((r, i) => `
      <div class="leader-row" data-portfolio="${escapeAttr(r.name)}" style="cursor:pointer;">
        <div class="leader-rank">${i + 1}</div>
        <div class="leader-name">${r.name}</div>
        <div class="leader-bar"><span style="width:${r.val}%;background:${r.val >= 80 ? '#22c55e' : r.val >= 60 ? '#f5a524' : '#ef4444'}"></span></div>
        <div class="leader-val">${r.val}%</div>
      </div>`).join("");
    $all("[data-portfolio]", $("#leaderboard")).forEach(row => row.addEventListener("click", () => openPortfolio(row.getAttribute("data-portfolio"))));

    // KPI deep-dive: one bespoke visualization per metric, not a generic dropdown+bar.
    renderKpiDeepDive(names, metricNames);
  }

  // ---------------------------------------------------------------
  // KPI Deep-Dive — per-metric visualizations, chosen by what the metric means
  // ---------------------------------------------------------------
  const METRIC_VIZ = {
    "Defect Leakage": { viz: "rankedBar", theme: "leak", caption: "Ranked lowest → highest. Lower leakage is better." },
    "Level of Automated Testing": { viz: "bullet", caption: "Bold bar = automation achieved. Pale bar = this portfolio's target." },
    "Customer Satisfaction": { viz: "leaderboard", caption: "Ranked satisfaction score — top 3 medalled." },
    "Risk-Based Test Coverage": { viz: "stacked", caption: "Coverage achieved vs. the gap still remaining to 100%." },
    "Defect Detection Rate": { viz: "ruler", markerShape: "circle", caption: "Each marker is a portfolio's detection rate; the tick is its own target." },
    "Test Execution Downtime": { viz: "rankedBar", theme: "time", caption: "Ranked fewest → most hours lost. Lower is better." },
    "Defect Turnaround Time": { viz: "ruler", markerShape: "diamond", caption: "Position on the day-scale; the tick marks the contractual target." },
    "Critical Test Milestone Delays": { viz: "heat", caption: "One tile per portfolio — darker means more milestone delays this month." },
  };
  function slug(s) { return String(s).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/(^-|-$)/g, ""); }

  function renderKpiDeepDive(names, metricNames) {
    const root = $("#kpi-deepdive");
    if (!metricNames.length) { root.innerHTML = `<p class="text-sub">No KPI data loaded for this month.</p>`; return; }

    root.innerHTML = metricNames.map(metric => {
      const cfg = METRIC_VIZ[metric] || { viz: "rankedBar", caption: "Ranked across all portfolios." };
      const span = (cfg.viz === "leaderboard" || cfg.viz === "heat" || cfg.viz === "ruler") ? 1 : 2;
      const id = "viz-" + slug(metric);
      return `
        <div class="card card-pad" style="grid-column:span ${span};">
          <div class="card-head"><div><div class="card-title">${metric}</div><div class="card-sub">${cfg.caption}</div></div></div>
          <div id="${id}"></div>
        </div>`;
    }).join("");

    metricNames.forEach(metric => {
      const cfg = METRIC_VIZ[metric] || { viz: "rankedBar" };
      const rows = names.map(n => {
        const m = getPortfolio(state.selectedMonth, n).kpi.find(x => x.name === metric);
        return { name: n, actual: m ? m.actual : null, rag: m ? m.rag : "", target: m ? m.target : null, unit: m ? m.unit : "", direction: m ? m.direction : "Higher" };
      }).filter(r => r.actual != null);
      const id = "viz-" + slug(metric);
      if (!rows.length) { $("#" + id).innerHTML = `<p class="text-sub" style="font-size:13px;">No data.</p>`; return; }

      if (cfg.viz === "rankedBar") renderVizRankedBar(id, metric, rows, cfg.theme);
      else if (cfg.viz === "bullet") renderVizBullet(id, metric, rows);
      else if (cfg.viz === "stacked") renderVizStacked(id, metric, rows);
      else if (cfg.viz === "leaderboard") renderVizLeaderboard(id, rows);
      else if (cfg.viz === "ruler") renderVizRuler(id, rows, cfg.markerShape);
      else if (cfg.viz === "heat") renderVizHeat(id, rows);
    });
  }

  function vizCanvas(id, heightPx) {
    const wrap = document.createElement("div");
    wrap.className = "chart-box";
    wrap.style.height = heightPx + "px";
    wrap.innerHTML = `<canvas></canvas>`;
    $("#" + id).appendChild(wrap);
    return wrap.querySelector("canvas").getContext("2d");
  }

  function renderVizRankedBar(id, metric, rows, theme) {
    const higherBetter = rows[0].direction === "Higher";
    const sorted = rows.slice().sort((a, b) => higherBetter ? b.actual - a.actual : a.actual - b.actual);
    const ctx = vizCanvas(id, Math.max(220, sorted.length * 26));
    const palette = theme === "time"
      ? sorted.map(r => RAG_COLOR[r.rag] || "#94a3b8")
      : sorted.map(r => RAG_COLOR[r.rag] || "#c9d1e0");
    makeChart("viz-" + slug(metric), ctx, {
      type: "bar",
      data: { labels: sorted.map(r => r.name), datasets: [{ data: sorted.map(r => r.actual), backgroundColor: palette, borderRadius: 5, barThickness: 14 }] },
      options: {
        indexAxis: "y",
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: (c) => `${c.raw}${sorted[c.dataIndex].unit}  (target ${sorted[c.dataIndex].target}${sorted[c.dataIndex].unit})` } }
        },
        scales: { x: { beginAtZero: true, grid: { color: "#f0f2f6" } }, y: { grid: { display: false } } },
        onClick: (evt, els) => { if (els.length) openPortfolio(sorted[els[0].index].name); }
      }
    });
  }

  function renderVizBullet(id, metric, rows) {
    const sorted = rows.slice().sort((a, b) => b.actual - a.actual);
    const ctx = vizCanvas(id, Math.max(220, sorted.length * 28));
    makeChart("viz-" + slug(metric), ctx, {
      type: "bar",
      data: {
        labels: sorted.map(r => r.name),
        datasets: [
          { label: "Target", data: sorted.map(r => r.target), backgroundColor: "#e6e9f0", barThickness: 20, borderRadius: 3, order: 2 },
          { label: "Actual", data: sorted.map(r => r.actual), backgroundColor: sorted.map(r => RAG_COLOR[r.rag] || "#0ea5a0"), barThickness: 9, borderRadius: 3, order: 1 },
        ]
      },
      options: {
        indexAxis: "y",
        maintainAspectRatio: false,
        datasets: { bar: { grouped: false } },
        plugins: {
          legend: { position: "bottom", labels: { boxWidth: 10, font: { size: 10.5 } } },
          tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${c.raw}${sorted[c.dataIndex].unit}` } }
        },
        scales: { x: { beginAtZero: true, grid: { color: "#f0f2f6" } }, y: { grid: { display: false } } },
        onClick: (evt, els) => { if (els.length) openPortfolio(sorted[els[0].index].name); }
      }
    });
  }

  function renderVizStacked(id, metric, rows) {
    const sorted = rows.slice().sort((a, b) => b.actual - a.actual);
    const ctx = vizCanvas(id, Math.max(220, sorted.length * 26));
    makeChart("viz-" + slug(metric), ctx, {
      type: "bar",
      data: {
        labels: sorted.map(r => r.name),
        datasets: [
          { label: "Covered", data: sorted.map(r => r.actual), backgroundColor: sorted.map(r => RAG_COLOR[r.rag] || "#0ea5a0"), stack: "s", barThickness: 14 },
          { label: "Gap", data: sorted.map(r => Math.max(0, 100 - r.actual)), backgroundColor: "#eef1f6", stack: "s", barThickness: 14 },
        ]
      },
      options: {
        indexAxis: "y",
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "bottom", labels: { boxWidth: 10, font: { size: 10.5 } } },
          tooltip: { callbacks: { label: (c) => `${c.dataset.label}: ${round1(c.raw)}%` } }
        },
        scales: { x: { stacked: true, max: 100, grid: { color: "#f0f2f6" } }, y: { stacked: true, grid: { display: false } } },
        onClick: (evt, els) => { if (els.length) openPortfolio(sorted[els[0].index].name); }
      }
    });
  }

  function renderVizLeaderboard(id, rows) {
    const sorted = rows.slice().sort((a, b) => b.actual - a.actual);
    const medals = ["🥇", "🥈", "🥉"];
    $("#" + id).innerHTML = sorted.map((r, i) => `
      <div class="leader-row" data-portfolio="${escapeAttr(r.name)}" style="cursor:pointer;">
        <div class="leader-rank" style="${i < 3 ? "background:transparent;font-size:15px;" : ""}">${i < 3 ? medals[i] : i + 1}</div>
        <div class="leader-name">${r.name}</div>
        <div class="leader-bar"><span style="width:${r.actual}%;background:${RAG_COLOR[r.rag] || '#0ea5a0'}"></span></div>
        <div class="leader-val">${r.actual}${r.unit}</div>
      </div>`).join("");
    $all("[data-portfolio]", $("#" + id)).forEach(row => row.addEventListener("click", () => openPortfolio(row.getAttribute("data-portfolio"))));
  }

  function renderVizRuler(id, rows, shape) {
    const max = Math.max(...rows.map(r => Math.max(r.actual, r.target || 0))) * 1.15 || 1;
    const sorted = rows.slice().sort((a, b) => a.direction === "Higher" ? b.actual - a.actual : a.actual - b.actual);
    $("#" + id).innerHTML = `
      <div style="display:flex;flex-direction:column;gap:9px;">
        ${sorted.map(r => {
          const posActual = Math.min(96, (r.actual / max) * 100);
          const posTarget = r.target != null ? Math.min(96, (r.target / max) * 100) : null;
          const markerHtml = shape === "diamond"
            ? `<div style="position:absolute;left:${posActual}%;top:50%;width:9px;height:9px;transform:translate(-50%,-50%) rotate(45deg);background:${RAG_COLOR[r.rag] || '#0ea5a0'};border:2px solid #fff;box-shadow:0 0 0 1px ${RAG_COLOR[r.rag] || '#0ea5a0'};"></div>`
            : `<div style="position:absolute;left:${posActual}%;top:50%;width:11px;height:11px;transform:translate(-50%,-50%);border-radius:50%;background:${RAG_COLOR[r.rag] || '#0ea5a0'};border:2px solid #fff;box-shadow:0 0 0 1px ${RAG_COLOR[r.rag] || '#0ea5a0'};"></div>`;
          return `
          <div style="display:flex;align-items:center;gap:10px;cursor:pointer;" data-portfolio="${escapeAttr(r.name)}">
            <div style="width:120px;font-size:12px;font-weight:600;flex:0 0 auto;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${r.name}</div>
            <div style="position:relative;flex:1;height:6px;background:#eef1f6;border-radius:6px;">
              ${posTarget != null ? `<div style="position:absolute;left:${posTarget}%;top:-4px;width:2px;height:14px;background:#98a2b3;"></div>` : ""}
              ${markerHtml}
            </div>
            <div style="width:52px;text-align:right;font-size:12px;font-weight:700;flex:0 0 auto;">${r.actual}${r.unit}</div>
          </div>`;
        }).join("")}
      </div>
      <div style="margin-top:8px;font-size:10.5px;color:var(--sub);">Grey tick = this portfolio's own target</div>
    `;
    $all("[data-portfolio]", $("#" + id)).forEach(row => row.addEventListener("click", () => openPortfolio(row.getAttribute("data-portfolio"))));
  }

  function renderVizHeat(id, rows) {
    const sorted = rows.slice().sort((a, b) => b.actual - a.actual);
    $("#" + id).innerHTML = `
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(64px,1fr));gap:8px;">
        ${sorted.map(r => {
          const bg = r.actual <= 0 ? "#e7f7ee" : r.actual === 1 ? "#fff3d6" : "#fde3e1";
          const fg = r.actual <= 0 ? "#067647" : r.actual === 1 ? "#b54708" : "#b42318";
          return `
          <div data-portfolio="${escapeAttr(r.name)}" style="cursor:pointer;border-radius:10px;background:${bg};padding:10px 6px;text-align:center;">
            <div style="font-size:18px;font-weight:800;color:${fg};">${r.actual}</div>
            <div style="font-size:9.5px;color:${fg};opacity:.85;margin-top:2px;line-height:1.25;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${escapeAttr(r.name)}">${r.name}</div>
          </div>`;
        }).join("")}
      </div>
    `;
    $all("[data-portfolio]", $("#" + id)).forEach(cell => cell.addEventListener("click", () => openPortfolio(cell.getAttribute("data-portfolio"))));
  }

  // ---------------------------------------------------------------
  // Empty state / navigation / topbar
  // ---------------------------------------------------------------
  function emptyStateHTML() {
    return `
      <div class="empty-state">
        <div class="icon-wrap">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M4 20h16"/></svg>
        </div>
        <h2>No data loaded yet</h2>
        <p>This dashboard reads directly from Excel workbooks on your computer — nothing is uploaded anywhere. Point it at your ATS_Data folder to get started.</p>
        <div class="steps">
          1. Ask each portfolio lead to fill in the <code>ATS_Portfolio_Template.xlsx</code> from the <code>/templates</code> folder.<br>
          2. Save each completed file into <code>ATS_Data/&lt;YYYY-MM&gt;/&lt;Portfolio Name&gt;.xlsx</code> — e.g. <code>ATS_Data/2026-09/Self Assessment.xlsx</code>.<br>
          3. Click <b>Load Data</b> below and select the <code>ATS_Data</code> folder.<br>
          4. Repeat each month — the dashboard will automatically show month-over-month comparisons.
        </div>
        <button class="btn btn-primary" id="btn-load-empty" style="margin-top:22px;">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M4 20h16"/></svg>
          Load Data Folder
        </button>
      </div>`;
  }
  function wireEmptyState() {
    const btn = $("#btn-load-empty");
    if (btn) btn.addEventListener("click", () => $("#folder-input").click());
  }

  function switchView(view) {
    state.currentView = view;
    $all(".nav-item").forEach(b => b.classList.toggle("active", b.getAttribute("data-view") === view));
    $all(".view").forEach(v => v.classList.remove("active"));
    $("#view-" + view).classList.add("active");
    const titles = {
      overview: ["Program Overview", "All portfolios · executive summary"],
      portfolio: ["Portfolio Detail", state.selectedPortfolio || "Select a portfolio"],
      trends: ["Trends & Comparison", "Month-over-month and cross-portfolio analysis"],
      tsr: ["TSR Register", "All Test Scope Requests across the program"],
    };
    $("#view-title").textContent = titles[view][0];
    $("#view-crumb").textContent = titles[view][1];
    renderCurrentView();
    window.scrollTo({ top: 0, behavior: "instant" in window ? "instant" : "auto" });
  }

  function renderCurrentView() {
    if (state.currentView === "overview") renderOverview();
    else if (state.currentView === "portfolio") renderPortfolioView();
    else if (state.currentView === "trends") renderTrendsView();
    else if (state.currentView === "tsr") renderTsrView();
    updateNavBadge();
  }

  function updateNavBadge() {
    const badge = $("#nav-tsr-count");
    if (!state.selectedMonth) { badge.textContent = ""; return; }
    const total = allTsrForMonth(state.selectedMonth).length;
    badge.textContent = total || "";
  }

  function populateMonthSelect() {
    const sel = $("#month-select");
    sel.innerHTML = state.months.map(m => `<option value="${m}" ${m === state.selectedMonth ? "selected" : ""}>${fmtMonth(m)}</option>`).join("");
    sel.disabled = !state.months.length;
  }

  function renderAll() {
    populateMonthSelect();
    updateDataStatus();
    renderCurrentView();
  }

  // ---------------------------------------------------------------
  // Present mode
  // ---------------------------------------------------------------
  function togglePresent(on) {
    document.body.classList.toggle("present-mode", on);
    $("#btn-present").style.display = on ? "none" : "";
    $("#btn-print").style.display = on ? "none" : "";
    $("#btn-present-exit").style.display = on ? "inline-flex" : "none";
  }

  // ---------------------------------------------------------------
  // Wiring
  // ---------------------------------------------------------------
  function wire() {
    $all(".nav-item[data-legacy]").forEach(btn => btn.addEventListener("click", () => { if (window.ATS && ATS.shell) ATS.shell.go(btn.getAttribute("data-view")); else switchView(btn.getAttribute("data-view")); }));

    $("#folder-input").addEventListener("change", (e) => {
      if (!e.target.files.length) return;
      ingestFiles(e.target.files);
      e.target.value = "";
    });
    $("#btn-load-sidebar").addEventListener("click", () => $("#folder-input").click());

    $("#month-select").addEventListener("change", (e) => { state.selectedMonth = e.target.value; renderAll(); });

    $all("#compare-toggle button").forEach(btn => btn.addEventListener("click", () => {
      state.compareOn = btn.getAttribute("data-cmp") === "on";
      $all("#compare-toggle button").forEach(b => b.classList.toggle("active", b === btn));
      renderCurrentView();
    }));

    $("#btn-present").addEventListener("click", () => togglePresent(true));
    $("#btn-present-exit").addEventListener("click", () => togglePresent(false));
    $("#btn-print").addEventListener("click", () => window.print());

    $("#modal-close").addEventListener("click", () => $("#modal-backdrop").classList.remove("open"));
    $("#modal-backdrop").addEventListener("click", (e) => { if (e.target.id === "modal-backdrop") $("#modal-backdrop").classList.remove("open"); });
  }

  // ---------------------------------------------------------------
  // Init
  // ---------------------------------------------------------------
  function init() {
    wire();
    const hadPersisted = loadPersisted();
    if (hadPersisted) { populateMonthSelect(); }
    renderCurrentView();
    if (!hadPersisted) {
      $("#view-overview").innerHTML = emptyStateHTML();
      wireEmptyState();
    }
  }

  document.addEventListener("DOMContentLoaded", init);
  window.ATSLegacy = { show: switchView, hasData: () => !!state.selectedMonth, rerender: renderCurrentView };
})();
