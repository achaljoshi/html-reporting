/* ============================================================
   Per-KPI page visuals. ATS.kpiPages[key](res, ctx) -> { left:html, draw(root) }
   ctx = { data, period, all, prefix, raid, poap }
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date, N = ATS.num, R = ATS.rag, K = ATS.kpi, U = ATS.ui, esc = ATS.dom.esc;
  const C = U.C;
  const P = (ATS.kpiPages = {});
  const pct = (v) => (v == null ? null : Math.round(v * 1000) / 10);
  const cv = (root, id) => root.querySelector("#" + id);
  const ragFill = (r) => R.color(r);
  const wk = (we) => D.fmt(we, false);
  const lineDS = (label, data, color, extra) => Object.assign({ label, data, borderColor: color, backgroundColor: color + "22", tension: 0.3, pointRadius: 3, borderWidth: 2, spanGaps: true }, extra || {});
  const tick = (cb) => ({ callback: cb, color: "#667085", font: { size: 11 } });
  const pctAxis = (max) => ({ grid: { color: C.grid }, beginAtZero: true, max: max || 100, ticks: tick((v) => v + "%") });

  // ---------------------------------------------------------------- coverage
  P.coverage = function (res, ctx) {
    const d = res.detail, lv = d.levels, p = ctx.prefix;
    const stats = `<div class="sx-stats">
      ${U.stat(N.int(d.totReqs), "Requirements")}${U.stat(N.int(d.totTcs), "Test cases")}
      ${U.stat(N.pct1(lv.high.pct), "High risk", U.deltaChip(res))}${U.stat(N.pct1(lv.med.pct), "Medium risk")}${U.stat(N.pct1(lv.low.pct), "Low risk")}</div>`;
    const left = `${stats}
      <div class="sx-grid2">${U.chartBox(p + "cov-a", 230, "Coverage % vs target", "by risk level")}${U.chartBox(p + "cov-b", 230, "Requirements vs test cases", "by risk level")}</div>
      <div class="sx-grid2">${U.chartBox(p + "cov-c", 220, "Coverage trend", "weekly")}${d.scope.length ? U.chartBox(p + "cov-d", 220, "Test case scope by topic", "planned test cases") : ""}</div>`;
    return {
      left,
      draw(root) {
        const L = [["High", lv.high], ["Medium", lv.med], ["Low", lv.low]];
        ATS.chart(p + "cov-a", cv(root, p + "cov-a"), { type: "bar", data: { labels: L.map((x) => x[0]), datasets: [
          { label: "Coverage %", data: L.map((x) => pct(x[1].pct)), backgroundColor: L.map((x) => ragFill(x[1].rag)), borderRadius: 6, maxBarThickness: 46 },
          { type: "line", label: "Target %", data: L.map((x) => pct(x[1].target)), borderColor: C.navy, backgroundColor: C.navy, pointStyle: "rectRot", pointRadius: 7, showLine: false } ] },
          options: Object.assign(U.base(), { scales: { x: U.axes().x, y: pctAxis(100) } }) });
        ATS.chart(p + "cov-b", cv(root, p + "cov-b"), { type: "bar", data: { labels: L.map((x) => x[0]), datasets: [
          { label: "Requirements", data: L.map((x) => x[1].reqs), backgroundColor: C.navy, borderRadius: 5 }, { label: "Test cases", data: L.map((x) => x[1].tcs), backgroundColor: C.teal, borderRadius: 5 } ] },
          options: Object.assign(U.base(), { scales: U.axes() }) });
        ATS.chart(p + "cov-c", cv(root, p + "cov-c"), { type: "line", data: { labels: d.series.map((s) => wk(s.we)), datasets: [lineDS("High", d.series.map((s) => pct(s.high)), C.red), lineDS("Medium", d.series.map((s) => pct(s.med)), C.amber), lineDS("Low", d.series.map((s) => pct(s.low)), C.slate)] },
          options: Object.assign(U.base(), { scales: { x: U.axes().x, y: pctAxis(100) } }) });
        if (d.scope.length) {
          const cols = [C.teal, C.navy, C.indigo, C.amber, C.sky, C.violet, C.slate, C.green, C.red];
          ATS.chart(p + "cov-d", cv(root, p + "cov-d"), { type: "doughnut", data: { labels: d.scope.map((s) => s.topic), datasets: [{ data: d.scope.map((s) => s.planned), backgroundColor: d.scope.map((_, i) => cols[i % cols.length]), borderWidth: 0 }] },
            options: Object.assign(U.base(), { cutout: "62%", plugins: { legend: { position: "right", labels: { boxWidth: 9, font: { size: 10 }, usePointStyle: true } } } }) });
        }
      },
    };
  };

  // ---------------------------------------------------------------- defect detection
  P.dde = function (res, ctx) {
    const d = res.detail, p = ctx.prefix;
    if (d.empty) return { left: U.noData("No defects raised yet"), draw() {} };
    const weeks = Object.keys(d.raisedByWeek).sort();
    const list = (res.detail.open || []).slice(0, 8);
    const left = `<div class="sx-stats">${U.stat(d.ft, "Found in testing", U.deltaChip(res))}${U.stat(d.fp, "Escaped to prod")}${U.stat(d.dde == null ? "—" : N.pct(d.dde), "Detection effectiveness")}${U.stat(d.perPeriod, "Raised this " + ctx.period.type)}${U.stat(N.pct1(d.rejPct), "Rejection rate")}</div>
      <div class="sx-grid2">${U.chartBox(p + "dd-a", 230, "Raised vs rejected", "by priority, this " + ctx.period.type)}${U.chartBox(p + "dd-b", 230, "Defects raised per week", "bars = new, line = cumulative")}</div>
      <div><div class="sx-ctitle">Open defects ${d.open.length ? "" : "— none"}</div>${d.open.length ? U.table(["ID", "Summary", "Priority", "Status", "Raised", "TCs blocked"], list.map((x) => [`<b>${esc(x.id)}</b>`, esc(x.summary), esc(x.priority), esc(x.status), D.fmt(x.raised, false), N.int(x.blocked)]), { wrap: true }) : ""}</div>`;
    return {
      left,
      draw(root) {
        ATS.chart(p + "dd-a", cv(root, p + "dd-a"), { type: "bar", data: { labels: d.byPri.map((x) => x.pri), datasets: [
          { label: "Raised", data: d.byPri.map((x) => x.raised), backgroundColor: C.teal, borderRadius: 6 }, { label: "Rejected", data: d.byPri.map((x) => x.rejected), backgroundColor: C.red, borderRadius: 6 } ] },
          options: Object.assign(U.base(), { scales: Object.assign(U.axes(), { y: { grid: { color: C.grid }, beginAtZero: true, ticks: { precision: 0, color: "#667085" } } }) }) });
        let cum = 0;
        const cumul = weeks.map((w) => (cum += d.raisedByWeek[w]));
        ATS.chart(p + "dd-b", cv(root, p + "dd-b"), { type: "bar", data: { labels: weeks.map(wk), datasets: [
          { type: "line", label: "Cumulative", data: cumul, borderColor: C.navy, backgroundColor: C.navy, tension: 0.25, pointRadius: 3, yAxisID: "y1" },
          { label: "Raised", data: weeks.map((w) => d.raisedByWeek[w]), backgroundColor: C.teal, borderRadius: 5 } ] },
          options: Object.assign(U.base(), { scales: { x: U.axes().x, y: { grid: { color: C.grid }, beginAtZero: true, ticks: { precision: 0 } }, y1: { position: "right", grid: { display: false }, beginAtZero: true, ticks: { precision: 0 } } } }) });
      },
    };
  };

  // ---------------------------------------------------------------- environment
  P.environment = function (res, ctx) {
    const d = res.detail, p = ctx.prefix, th = ctx.data.th;
    const left = `<div class="sx-stats">${U.stat(N.pct1(d.avail), "Availability", U.deltaChip(res))}${U.stat(N.round1(d.down) + "h", "Downtime")}${U.stat(N.round1(d.req - d.down) + "h", "Uptime")}${U.stat(N.round1(d.req) + "h", "Required hours")}${U.stat(d.days, "Test days")}</div>
      <div class="sx-grid2">${U.chartBox(p + "env-a", 230, "Availability vs target", "% of required hours")}${U.chartBox(p + "env-b", 230, "Downtime hours per week", "threshold lines = weekly limits")}</div>`;
    return {
      left,
      draw(root) {
        const target = d.req > 0 ? 1 - d.greenAt / d.req : 1;
        ATS.chart(p + "env-a", cv(root, p + "env-a"), { type: "bar", data: { labels: ["Availability %", "Target %"], datasets: [
          { label: "Uptime", data: [pct(d.avail), pct(target)], backgroundColor: C.teal, borderRadius: 4 }, { label: "Downtime", data: [pct(1 - d.avail), pct(1 - target)], backgroundColor: C.red, borderRadius: 4 } ] },
          options: Object.assign(U.base(), { scales: { x: Object.assign(U.axes().x, { stacked: true }), y: { stacked: true, max: 100, grid: { color: C.grid }, ticks: tick((v) => v + "%") } } }) });
        const s = d.series.slice(-12);
        ATS.chart(p + "env-b", cv(root, p + "env-b"), { type: "bar", data: { labels: s.map((x) => wk(x.we)), datasets: [
          { label: "Downtime (h)", data: s.map((x) => x.down), backgroundColor: s.map((x) => (x.down <= th.dtWGreen ? C.teal : x.down <= th.dtWAmber ? C.amber : C.red)), borderRadius: 5 },
          { type: "line", label: "Green limit", data: s.map(() => th.dtWGreen), borderColor: C.green, borderDash: [5, 4], pointRadius: 0, borderWidth: 1.5 },
          { type: "line", label: "Amber limit", data: s.map(() => th.dtWAmber), borderColor: C.amber, borderDash: [5, 4], pointRadius: 0, borderWidth: 1.5 } ] },
          options: Object.assign(U.base(), { scales: U.axes() }) });
      },
    };
  };

  // ---------------------------------------------------------------- aging
  P.aging = function (res, ctx) {
    const d = res.detail, p = ctx.prefix;
    if (res.empty) return { left: U.noData("No defects to age"), draw() {} };
    const open = d.open.slice().sort((a, b) => b.days - a.days).slice(0, 8);
    const left = `<div class="sx-stats">${U.stat(N.pct(d.compliance), "Within SLA (" + d.sla + "d)", U.deltaChip(res))}${U.stat(d.list.length, "Defects in scope")}${U.stat(d.open.length, "Open")}${U.stat(d.avg == null ? "—" : d.avg.toFixed(1) + "d", "Avg turnaround")}${U.stat(d.openOver10.length, "Open > 10 days")}</div>
      <div class="sx-grid2">${U.chartBox(p + "ag-a", 230, "Aging buckets", "working days")}${U.chartBox(p + "ag-b", 230, "Aging per defect", "working days vs SLA")}</div>
      ${open.length ? `<div><div class="sx-ctitle">Oldest open defects</div>${U.table(["ID", "Summary", "Priority", "Age (WD)", "Owner"], open.map((x) => [`<b>${esc(x.d.id)}</b>`, esc(x.d.summary), esc(x.d.priority), `<b style="color:${x.days > 10 ? C.red : "inherit"}">${x.days}</b>`, esc(x.d.owner)]), { wrap: true })}</div>` : ""}`;
    return {
      left,
      draw(root) {
        const cols = [C.green, C.teal, C.amber, C.red];
        ATS.chart(p + "ag-a", cv(root, p + "ag-a"), { type: "bar", data: { labels: K.BUCKETS, datasets: [{ label: "Defects", data: K.BUCKETS.map((b) => d.counts[b]), backgroundColor: cols, borderRadius: 6, maxBarThickness: 52 }] },
          options: Object.assign(U.base(), { plugins: { legend: { display: false } }, scales: Object.assign(U.axes(), { y: { grid: { color: C.grid }, beginAtZero: true, ticks: { precision: 0 } } }) }) });
        const l = d.list.slice().sort((a, b) => (a.d.raised < b.d.raised ? -1 : 1)).slice(-14);
        ATS.chart(p + "ag-b", cv(root, p + "ag-b"), { type: "bar", data: { labels: l.map((x) => x.d.id.replace(/^.*-/, "#")), datasets: [
          { label: "Aging (WD)", data: l.map((x) => x.days), backgroundColor: l.map((x) => (x.days <= d.sla ? C.teal : x.days <= 10 ? C.amber : C.red)), borderRadius: 4 },
          { type: "line", label: "SLA", data: l.map(() => d.sla), borderColor: C.navy, borderDash: [5, 4], pointRadius: 0, borderWidth: 1.5 },
          { type: "line", label: "Average", data: l.map(() => (d.avg == null ? null : d.avg)), borderColor: C.violet, pointRadius: 0, borderWidth: 1.5 } ] },
          options: Object.assign(U.base(), { scales: U.axes() }) });
      },
    };
  };

  // ---------------------------------------------------------------- milestones
  P.milestones = function (res, ctx) {
    const d = res.detail, p = ctx.prefix;
    if (res.empty) return { left: U.noData("No milestones"), draw() {} };
    const ev = d.ev;
    const dates = ev.map((e) => e.m.due).concat(ev.filter((e) => e.m.completed).map((e) => e.m.completed));
    const mn = dates.reduce((a, b) => (a < b ? a : b)), mx = dates.reduce((a, b) => (a > b ? a : b));
    const span = Math.max(1, D.diff(mn, mx)), pos = (iso) => Math.max(1, Math.min(99, (D.diff(mn, iso) / span) * 100));
    const asOf = ctx.all.asOf, todayPos = pos(asOf < mn ? mn : asOf > mx ? mx : asOf);
    const ribbon = `<div class="sx-ribbon"><div class="line"></div><div class="done" style="width:${todayPos}%"></div><div class="today" style="left:${todayPos}%"></div>
      ${ev.map((e, i) => `<div class="sx-ms ${i % 2 ? "up" : ""}" style="left:${pos(e.m.due)}%;color:${ragFill(e.rag)}" title="${esc(e.m.name)} — due ${D.fmt(e.m.due)} — ${esc(e.status)}${e.delay ? " (" + e.delay + "d)" : ""}"><b>${D.fmt(e.m.due, false)}</b><i></i><span>${esc(e.m.name.length > 28 ? e.m.name.slice(0, 27) + "…" : e.m.name)}</span></div>`).join("")}</div>`;
    const left = `<div class="sx-stats">${U.stat(N.pct(d.adherence), "Adherence", U.deltaChip(res))}${U.stat(d.completed + d.late, "Completed")}${U.stat(d.onSched, "On schedule")}${U.stat(d.overdue, "Overdue")}${U.stat(d.critDelay + "d", "Critical delay")}</div>
      <div><div class="sx-ctitle">Milestone timeline <span class="sx-csub">· colour = RAG, line = today</span></div>${ribbon}</div>
      ${U.table(["Milestone", "Due", "Forecast", "Completed", "Status", "RAG"], ev.map((e) => [esc(e.m.name) + (e.m.critical ? ' <span class="chip red" style="padding:0 6px">critical</span>' : ""), D.fmt(e.m.due, false), e.m.forecast ? D.fmt(e.m.forecast, false) : "—", e.m.completed ? D.fmt(e.m.completed, false) : "—", esc(e.status) + (e.delay ? ` (${e.delay}d)` : ""), R.chip(e.rag)]), { wrap: true })}`;
    return { left, draw() {} };
  };

  // ---------------------------------------------------------------- TSR
  P.tsr = function (res, ctx) {
    const d = res.detail, p = ctx.prefix;
    const items = d.items || [];
    const left = `<div class="sx-stats">${U.stat(items.length, "TSRs in scope")}${U.stat(d.recvIn || 0, "Received this " + ctx.period.type)}${U.stat(d.over || 0, "Over " + d.sla + "-day limit")}${U.stat(d.avg == null ? "—" : d.avg.toFixed(1) + "d", "Avg working days")}</div>
      ${items.length ? U.chartBox(p + "tsr-a", 230, "Working days per TSR", "limit = " + d.sla + " days (hold days excluded)") : '<div class="sx-callout green">No TSRs required assessment in this period.</div>'}
      ${items.length ? U.table(["TSR", "Description", "Received", "Returned", "Hold", "Working days", "SLA"], items.map((x) => [`<b>${esc(x.t.ref)}</b>`, esc(x.t.desc), D.fmt(x.t.received, false), x.open ? "Open" : D.fmt(x.t.returned, false), x.t.hold || 0, `<b>${x.wd}</b>`, x.over ? R.chip("Red", "Over") : R.chip("Green", "Within")]), { wrap: true }) : ""}`;
    return {
      left,
      draw(root) {
        if (!items.length) return;
        ATS.chart(p + "tsr-a", cv(root, p + "tsr-a"), { type: "bar", data: { labels: items.map((x) => x.t.ref.replace(/^TSR_?/, "").slice(-14)), datasets: [
          { label: "Working days", data: items.map((x) => x.wd), backgroundColor: items.map((x) => (x.over ? C.red : C.teal)), borderRadius: 5, maxBarThickness: 54 },
          { type: "line", label: "SLA limit", data: items.map(() => d.sla), borderColor: C.navy, borderDash: [6, 4], pointRadius: 0, borderWidth: 1.5 } ] },
          options: Object.assign(U.base(), { scales: U.axes() }) });
      },
    };
  };

  // ---------------------------------------------------------------- leakage
  P.leakage = function (res, ctx) {
    const d = res.detail, p = ctx.prefix;
    const banner = res.notInProd ? `<div class="sx-callout info"><b>Not in production.</b> The solution has not been deployed yet, so there is no leakage data for ${esc(d.quarter)}. This KPI activates automatically once a <i>Production Go-Live Date</i> is set in Config or a defect is logged with Found In = Production.</div>` : "";
    const left = `${banner}<div class="sx-stats">${U.stat(d.counts[0], "P1 in production")}${U.stat(d.counts[1], "P2 in production")}${U.stat(d.counts[2], "P3 in production")}${U.stat(esc(d.quarter), "Reporting quarter")}</div>${U.chartBox(p + "lk-a", 220, "Defects found in production", d.quarter)}`;
    return {
      left,
      draw(root) {
        ATS.chart(p + "lk-a", cv(root, p + "lk-a"), { type: "bar", data: { labels: K.PRIORITIES, datasets: [{ label: "Defects", data: d.counts, backgroundColor: [C.red, C.amber, C.slate], borderRadius: 6, maxBarThickness: 60 }] },
          options: Object.assign(U.base(), { plugins: { legend: { display: false } }, scales: Object.assign(U.axes(), { y: { grid: { color: C.grid }, beginAtZero: true, suggestedMax: 5, ticks: { precision: 0 } } }) }) });
      },
    };
  };

  // ---------------------------------------------------------------- CSAT
  P.csat = function (res, ctx) {
    const d = res.detail, p = ctx.prefix;
    if (!d.scores || !d.scores.length) return { left: `<div class="sx-callout info"><b>Not requested yet.</b> Enter a <i>CSAT Score (1-5)</i> in the Weekly_Snapshot sheet on the week it is collected; the KPI and its RAG then calculate automatically.</div>`, draw() {} };
    const left = `<div class="sx-stats">${U.stat(d.cur.csat.toFixed(1) + " / 5", "Current score", U.deltaChip(res))}${U.stat(d.prev ? d.prev.csat.toFixed(1) : "—", "Previous score")}${U.stat(d.pass, "Pass mark")}${U.stat(d.cur.csatResp || "—", "Responses")}</div>${U.chartBox(p + "cs-a", 230, "Customer satisfaction", "previous vs current")}`;
    return {
      left,
      draw(root) {
        const labels = d.scores.map((s) => wk(s.we));
        ATS.chart(p + "cs-a", cv(root, p + "cs-a"), { type: "bar", data: { labels, datasets: [{ label: "Score", data: d.scores.map((s) => s.score), backgroundColor: d.scores.map((s) => (s.score >= d.pass ? C.teal : C.red)), borderRadius: 6, maxBarThickness: 60 }, { type: "line", label: "Pass mark", data: d.scores.map(() => d.pass), borderColor: C.navy, borderDash: [5, 4], pointRadius: 0 }] },
          options: Object.assign(U.base(), { scales: { x: U.axes().x, y: { min: 0, max: 5, grid: { color: C.grid } } } }) });
      },
    };
  };

  // ---------------------------------------------------------------- automation
  P.automation = function (res, ctx) {
    const d = res.detail, p = ctx.prefix;
    const left = `<div class="sx-stats">${U.stat(N.pct1(d.td), "Test data automation", U.deltaChip(res))}${U.stat(N.pct1(d.tc), "Test case automation")}${U.stat(N.int(d.tdDone) + " / " + N.int(d.tdTotal), "Data setups automated")}${U.stat(N.int(d.tcDone) + " / " + N.int(d.tcTotal), "Test cases automated")}${U.stat(N.pct(d.target), "Target")}</div>
      <div class="sx-grid2">${U.chartBox(p + "au-a", 220, "Automation coverage", "automated vs manual")}${U.chartBox(p + "au-b", 220, "Automation trend", "weekly")}</div>`;
    return {
      left,
      draw(root) {
        const cats = [["Test Data Coverage", d.td], ["Test Case Coverage", d.tc]];
        ATS.chart(p + "au-a", cv(root, p + "au-a"), { type: "bar", data: { labels: cats.map((c) => c[0]), datasets: [
          { label: "Automated", data: cats.map((c) => pct(c[1]) || 0), backgroundColor: C.teal, borderRadius: 4 }, { label: "Not automated", data: cats.map((c) => (c[1] == null ? 0 : pct(1 - c[1]))), backgroundColor: "#e4e7ec", borderRadius: 4 } ] },
          options: Object.assign(U.base(), { indexAxis: "y", scales: { x: { stacked: true, max: 100, grid: { color: C.grid }, ticks: tick((v) => v + "%") }, y: { stacked: true, grid: { display: false } } } }) });
        ATS.chart(p + "au-b", cv(root, p + "au-b"), { type: "line", data: { labels: d.series.map((s) => wk(s.we)), datasets: [lineDS("Test data", d.series.map((s) => pct(s.td)), C.teal), lineDS("Test cases", d.series.map((s) => pct(s.tc)), C.indigo)] },
          options: Object.assign(U.base(), { scales: { x: U.axes().x, y: pctAxis(100) } }) });
      },
    };
  };

  // ---------------------------------------------------------------- commercial
  P.commercial = function (res, ctx) {
    const d = res.detail, p = ctx.prefix;
    const left = `<div class="sx-stats">${U.stat(N.gbpK(d.spend), "Spend to date", U.deltaChip(res))}${U.stat(N.gbpK(d.forecast), "Forecast at completion")}${U.stat(N.gbpK(d.budget), "Total budget")}${U.stat(N.pct(d.spendPct), "Budget consumed")}</div>${U.chartBox(p + "co-a", 250, "Spend, forecast & budget", "cumulative £")}
      ${ctx.data.config.tsrRef ? `<dl class="sx-kv"><dt>Project</dt><dd>${esc(ctx.data.config.project || ctx.data.config.programName)}</dd><dt>TSR</dt><dd>${esc(ctx.data.config.tsrRef)}</dd><dt>Start</dt><dd>${D.fmt(ctx.data.config.startDate)}</dd><dt>End</dt><dd>${D.fmt(ctx.data.config.endDate)}</dd></dl>` : ""}`;
    return {
      left,
      draw(root) {
        const s = d.series;
        ATS.chart(p + "co-a", cv(root, p + "co-a"), { type: "line", data: { labels: s.map((x) => wk(x.we)), datasets: [
          lineDS("Cumulative spend", s.map((x) => x.spend), C.teal, { fill: true }), lineDS("Forecast at completion", s.map((x) => x.forecast), C.indigo, { borderDash: [6, 4], backgroundColor: "transparent" }),
          lineDS("Budget", s.map(() => d.budget), C.red, { borderDash: [2, 3], pointRadius: 0, backgroundColor: "transparent" }) ] },
          options: Object.assign(U.base(), { scales: { x: U.axes().x, y: { grid: { color: C.grid }, ticks: tick((v) => "£" + Math.round(v / 1000) + "k") } } }) });
      },
    };
  };

  // ---------------------------------------------------------------- resource
  P.resource = function (res, ctx) {
    const d = res.detail, p = ctx.prefix;
    const left = `<div class="sx-stats">${U.stat(d.act + " / " + d.plan, "FTE in place vs plan", U.deltaChip(res))}${U.stat(N.pct(d.ratio), "Resourced")}${U.stat(d.roster.length, "Active roles")}${U.stat(d.planned.length, "Planned hires")}</div>
      <div class="sx-grid2">${U.chartBox(p + "rs-a", 220, "FTE planned vs actual", "weekly")}${U.chartBox(p + "rs-b", 220, "Resource mix", "by organisation (FTE)")}</div>
      ${U.table(["Role", "Name", "Organisation", "Allocation", "Status"], d.roster.concat(d.planned).slice(0, 14).map((x) => [esc(x.role), esc(x.name || "TBC"), esc(x.org), x.alloc == null ? "—" : N.pct(x.alloc), esc(x.status)]), { wrap: true })}`;
    return {
      left,
      draw(root) {
        ATS.chart(p + "rs-a", cv(root, p + "rs-a"), { type: "line", data: { labels: d.series.map((x) => wk(x.we)), datasets: [lineDS("Planned", d.series.map((x) => x.plan), C.navy, { stepped: true, borderDash: [5, 4], backgroundColor: "transparent" }), lineDS("Actual", d.series.map((x) => x.act), C.teal, { stepped: true, fill: true })] },
          options: Object.assign(U.base(), { scales: Object.assign(U.axes(), { y: { grid: { color: C.grid }, beginAtZero: true } }) }) });
        const orgs = Object.keys(d.orgs);
        ATS.chart(p + "rs-b", cv(root, p + "rs-b"), { type: "doughnut", data: { labels: orgs, datasets: [{ data: orgs.map((o) => N.round1(d.orgs[o])), backgroundColor: [C.teal, C.navy, C.indigo, C.amber, C.sky], borderWidth: 0 }] },
          options: Object.assign(U.base(), { cutout: "62%", plugins: { legend: { position: "right", labels: { boxWidth: 9, usePointStyle: true } } } }) });
      },
    };
  };

  P.demand = function (res, ctx) {
    const d = res.detail, p = ctx.prefix;
    const rows = d.rows || [];
    const left = `<div class="sx-stats">${U.stat(N.round1(d.maxGap) + " FTE", "Max gap (3 months)", U.deltaChip(res))}${U.stat(d.peak ? d.peak.planned + " FTE" : "—", "Peak demand")}${U.stat(d.upcoming.length, "Months visible")}</div>${U.chartBox(p + "dm-a", 250, "Demand pipeline forecast", "planned vs available FTE")}`;
    return {
      left,
      draw(root) {
        ATS.chart(p + "dm-a", cv(root, p + "dm-a"), { type: "bar", data: { labels: rows.map((r) => D.fmtMonth(r.month.slice(0, 7), false)), datasets: [
          { label: "Planned demand (FTE)", data: rows.map((r) => r.planned), backgroundColor: rows.map((r) => ((r.planned || 0) > (r.available == null ? r.planned : r.available) ? C.amber : C.teal)), borderRadius: 5 },
          { type: "line", label: "Available (FTE)", data: rows.map((r) => r.available), borderColor: C.navy, backgroundColor: C.navy, stepped: true, pointRadius: 3 } ] },
          options: Object.assign(U.base(), { scales: U.axes() }) });
      },
    };
  };

  // ---------------------------------------------------------------- execution
  P.execution = function (res, ctx) {
    const d = res.detail, p = ctx.prefix, t = d.total, pt = d.prevTotal;
    const dx = t.executed - pt.executed;
    const left = `<div class="sx-stats">${U.stat(N.int(t.planned), "Planned")}${U.stat(N.int(t.executed), "Executed", `<span class="sx-delta ${dx > 0 ? "good" : "neu"}">+${N.int(dx)} this ${ctx.period.type}</span>`)}${U.stat(N.int(t.passed), "Passed")}${U.stat(N.int(t.failed), "Failed")}${U.stat(N.int(t.blocked), "Blocked")}${U.stat(N.pct(d.passRate), "Pass rate")}</div>
      <div class="sx-grid2">${U.chartBox(p + "ex-a", 250, "Execution burn-up", "cumulative vs planned")}${U.chartBox(p + "ex-b", 250, "Weekly throughput", "test cases executed per week")}</div>
      ${U.table(["Topic", "Phase", "Planned", "Designed", "Passed", "Failed", "Blocked", "% executed"], d.rows.map((r) => [`<b>${esc(r.topic)}</b>`, esc(r.phase), N.int(r.planned), N.pct(r.planned ? r.designed / r.planned : null), N.int(r.passed), N.int(r.failed), N.int(r.blocked), N.pct(r.planned ? (r.passed + r.failed) / r.planned : null)]), { wrap: true })}`;
    return {
      left,
      draw(root) {
        const s = d.series;
        ATS.chart(p + "ex-a", cv(root, p + "ex-a"), { type: "bar", data: { labels: s.map((x) => wk(x.we)), datasets: [
          { label: "Passed", data: s.map((x) => x.passed), backgroundColor: C.teal, stack: "e", borderRadius: 2 }, { label: "Failed", data: s.map((x) => x.failed), backgroundColor: C.red, stack: "e", borderRadius: 2 },
          { label: "Blocked", data: s.map((x) => x.blocked), backgroundColor: C.amber, stack: "e", borderRadius: 2 },
          { type: "line", label: "Planned", data: s.map((x) => x.planned), borderColor: C.navy, borderDash: [6, 4], pointRadius: 0, borderWidth: 2 } ] },
          options: Object.assign(U.base(), { scales: { x: Object.assign(U.axes().x, { stacked: true }), y: { stacked: true, grid: { color: C.grid }, beginAtZero: true } } }) });
        const thr = s.map((x, i) => (i ? x.executed - s[i - 1].executed : x.executed));
        ATS.chart(p + "ex-b", cv(root, p + "ex-b"), { type: "bar", data: { labels: s.map((x) => wk(x.we)), datasets: [{ label: "Executed this week", data: thr, backgroundColor: C.indigo, borderRadius: 5 }] },
          options: Object.assign(U.base(), { plugins: { legend: { display: false } }, scales: U.axes() }) });
      },
    };
  };

  // ---------------------------------------------------------------- readiness
  P.readiness = function (res, ctx) {
    const d = res.detail;
    const areas = K.READINESS_AREAS || [];
    const cell = (a) => `<td><span class="sx-cellrag" style="background:${ragFill(a.rag)}" title="${esc(a.note || "No note")}">${U.ragLetter(a.rag)}</span></td>`;
    const left = `<div class="table-wrap"><table class="sx-matrix"><thead><tr><th style="text-align:left">Topic</th><th>SIT start</th><th>Overall</th>${areas.map((a) => `<th>${esc(a.label)}</th>`).join("")}</tr></thead><tbody>
      ${d.rows.map((r) => `<tr><td class="topic">${esc(r.topic)}</td><td>${r.sitStart ? D.fmt(r.sitStart, false) : "—"}</td><td><span class="sx-cellrag" style="background:${ragFill(r.overall)}">${U.ragLetter(r.overall)}</span></td>${areas.map((a) => cell(r.areas[a.key] || { rag: "" })).join("")}</tr>`).join("")}</tbody></table></div>
      <div class="sx-ctitle">Notes <span class="sx-csub">· hover a dot for the note</span></div>
      ${d.rows.map((r) => `<div class="sx-panel" style="margin-bottom:8px"><h5>${esc(r.topic)} ${R.chip(r.overall)}</h5>${areas.map((a) => (r.areas[a.key] && r.areas[a.key].note ? `<p style="margin-bottom:4px"><b>${esc(a.label)}:</b> ${esc(r.areas[a.key].note)}</p>` : "")).join("")}${r.roadToGreen ? `<p style="margin-top:6px;color:#0a7c78"><b>Road to Green:</b> ${esc(r.roadToGreen)}</p>` : ""}</div>`).join("")}`;
    return { left, draw() {} };
  };

  // ---------------------------------------------------------------- RAID
  P.raid = function (res, ctx) {
    const s = res.detail, p = ctx.prefix;
    if (res.empty) return { left: U.noData("RAID log not loaded"), draw() {} };
    const left = `<div class="sx-stats">${U.stat(s.openRisks, "Open risks")}${U.stat(s.openIssues, "Open issues")}${U.stat(s.openDeps, "Open dependencies")}${U.stat(s.unconfirmedAssumptions, "Unconfirmed assumptions")}${U.stat(s.decisionsPending, "Pending decisions")}${U.stat(s.veryHigh + s.high, "High / Very High")}</div>
      <div class="sx-grid2"><div><div class="sx-ctitle">Risk heat-map <span class="sx-csub">· open risks, click a cell</span></div><div id="${p}raid-heat">${ATS.raid.heatmap(s.risksOpen)}</div></div>
      <div><div class="sx-ctitle">Open risks by rating</div>${ATS.raid.bandBars(s.riskBands)}<div class="sx-ctitle" style="margin-top:14px">Open issues by rating</div>${ATS.raid.bandBars(s.issueBands)}</div></div>
      <div id="${p}raid-list"><div class="sx-ctitle">Top risks</div>${ATS.raid.riskTable(s.topRisks)}</div>
      <div><div class="sx-ctitle">Open issues</div>${ATS.raid.issueTable(s.topIssues)}</div>
      <div><div class="sx-ctitle">Dependencies needed in the next 30 days</div>${ATS.raid.depTable(s.depsDue.slice(0, 8))}</div>
      <div><div class="sx-ctitle">Decisions <span class="sx-csub">· ${s.decisionsPending} pending · ${s.decisionsInPeriod} taken in this period · ${s.decisionsTotal} logged</span></div>${ATS.raid.decisionTable(s.pendingDecisions.concat(s.recentDecisions).slice(0, 8), { why: true })}</div>`;
    return {
      left,
      draw(root) {
        const heat = root.querySelector("#" + p + "raid-heat"), list = root.querySelector("#" + p + "raid-list");
        if (!heat) return;
        heat.addEventListener("click", (e) => {
          const b = e.target.closest(".sx-heat-cell"); if (!b) return;
          const l = +b.dataset.l, i = +b.dataset.i;
          const sel = s.risksOpen.filter((r) => r.l === l && r.i === i);
          list.innerHTML = `<div class="sx-ctitle">Risks at likelihood ${l} × impact ${i} (${sel.length}) <button class="btn btn-outline btn-sm" id="${p}raid-reset" style="margin-left:8px">Show top risks</button></div>${ATS.raid.riskTable(sel, { titleLen: 160, mitLen: 220 })}`;
          const rb = list.querySelector("#" + p + "raid-reset");
          if (rb) rb.onclick = () => { list.innerHTML = `<div class="sx-ctitle">Top risks</div>${ATS.raid.riskTable(s.topRisks)}`; };
        });
      },
    };
  };
})();
