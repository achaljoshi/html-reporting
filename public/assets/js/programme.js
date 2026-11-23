/* ============================================================
   Programme Overview — every portfolio side by side, derived from each portfolio's
   own weekly workbook (same design for all). Click anything to drill into that
   portfolio's Monthly Council / Weekly Report.
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date, N = ATS.num, R = ATS.rag, K = ATS.kpi, U = ATS.ui, esc = ATS.dom.esc;
  const C = U.C;
  const P = (ATS.programme = {});
  const st = { mode: "month", key: null, tab: "heatmap", kpi: "execution", reg: "tsr", q: "", sort: "severity" };
  let root = null, last = null;

  const SHORT = { coverage: "Coverage", leakage: "Leakage", dde: "Detection", aging: "Aging", automation: "Automation", milestones: "Milestones", environment: "Environment", tsr: "TSR", csat: "CSAT", resource: "Resources", demand: "Demand", commercial: "Commercial", raid: "RAID", execution: "Execution", readiness: "Readiness" };
  const RANK = { Grey: 0, Green: 1, Amber: 2, Red: 3 };

  P.mount = (el) => { root = el; };
  P.root = () => root;
  P.last = () => last;

  function portfolios() { return ATS.store.order.map((n) => Object.assign({ name: n }, ATS.store.portfolios[n])).filter((p) => p.kpi); }

  function periods(ps) {
    const months = new Set(), weeks = new Set();
    ps.forEach((p) => { K.months(p.kpi).forEach((m) => months.add(m)); K.weeks(p.kpi).forEach((w) => weeks.add(w)); });
    return { month: Array.from(months).sort(), week: Array.from(weeks).sort() };
  }

  function compute(ps, mode, key) {
    const period = mode === "month" ? K.monthPeriod(key) : K.weekPeriod(key);
    const rows = ps.map((p) => ({ name: p.name, p, all: K.computeAll(p.kpi, period, { raid: p.raid }) }));
    return { rows, period };
  }

  const kpi = (r, k) => r.all.byKey[k];
  const sev = (r) => RANK[r.all.overall.rag] * 100 + r.all.scorecard.filter((x) => x.rag === "Red").length * 10 + r.all.scorecard.filter((x) => x.rag === "Amber").length;

  function aggregate(rows) {
    const a = { n: rows.length, rag: { Green: 0, Amber: 0, Red: 0, Grey: 0 }, openDef: 0, tsrOver: 0, msOver: 0, spend: 0, budget: 0, exec: 0, planned: 0, raidHigh: 0 };
    rows.forEach((r) => {
      a.rag[r.all.overall.rag]++;
      const dd = kpi(r, "dde"), ts = kpi(r, "tsr"), ms = kpi(r, "milestones"), co = kpi(r, "commercial"), ex = kpi(r, "execution"), rd = kpi(r, "raid");
      if (!dd.empty && dd.detail.open) a.openDef += dd.detail.open.length;
      if (ts.detail && ts.detail.over) a.tsrOver += ts.detail.over;
      if (!ms.empty) a.msOver += ms.detail.overdue;
      if (!co.empty) { a.spend += co.detail.spend || 0; a.budget += co.detail.budget || 0; }
      if (!ex.empty) { a.exec += ex.detail.total.executed; a.planned += ex.detail.total.planned; }
      if (!rd.empty) a.raidHigh += rd.detail.veryHigh + rd.detail.high;
    });
    return a;
  }

  function open(name, page) {
    ATS.setPortfolio(name, true);
    const mod = st.mode === "month" ? ATS.monthly : ATS.weekly;
    mod.state.key = st.key; mod.state.page = page || (st.mode === "month" ? "scorecard" : "glance");
    ATS.shell.go(st.mode === "month" ? "monthly" : "weekly");
    if (page && mod.state.page !== page) {
      const m = K.KPI_META.find((x) => x.key === page);
      ATS.toast(`${m ? m.name : page} is a ${m ? m.cadence.toLowerCase() : "monthly"} measure and is not part of the Weekly Report — open the Monthly Council to see it.`);
    }
  }

  // ---------------------------------------------------------------- views
  function heatmap(rows) {
    const cols = K.KPI_META;
    const sorted = rows.slice().sort((a, b) => (st.sort === "name" ? a.name.localeCompare(b.name) : sev(b) - sev(a) || a.name.localeCompare(b.name)));
    return `<div class="table-wrap"><table class="sx-matrix sx-hm"><thead><tr><th style="text-align:left">Portfolio</th><th>Overall</th>${cols.map((c) => `<th class="sx-vh" title="${esc(c.name)}"><span>${SHORT[c.key]}</span></th>`).join("")}</tr></thead><tbody>
      ${sorted.map((r) => `<tr><td class="topic"><a data-open="${esc(r.name)}" style="cursor:pointer">${esc(r.name)}</a></td>
        <td><span class="sx-cellrag" style="background:${R.color(r.all.overall.rag)}" title="Overall ${esc(r.all.overall.rag)}${r.all.overall.declared ? " (declared)" : ""}">${U.ragLetter(r.all.overall.rag)}</span></td>
        ${cols.map((c) => { const x = kpi(r, c.key); return `<td><span class="sx-cellrag" data-open="${esc(r.name)}" data-page="${c.key}" style="background:${R.color(x.rag)};cursor:pointer" title="${esc(r.name + " — " + c.name + ": " + x.position)}">${U.ragLetter(x.rag)}</span></td>`; }).join("")}</tr>`).join("")}
      </tbody></table></div>
      <p class="sx-empty-note">One row per portfolio, one column per KPI — each dot recalculated from that portfolio's own weekly workbook as of the end of the selected period. Click a dot to open that KPI page for that portfolio.</p>`;
  }

  function cards(rows) {
    const sorted = rows.slice().sort((a, b) => sev(b) - sev(a) || a.name.localeCompare(b.name));
    return `<div class="sx-pgrid">${sorted.map((r) => {
      const sc = r.all.scorecard, g = sc.filter((x) => x.rag === "Green").length, a = sc.filter((x) => x.rag === "Amber").length, rd = sc.filter((x) => x.rag === "Red").length, n = g + a + rd || 1;
      const ex = kpi(r, "execution"), dd = kpi(r, "dde"), en = kpi(r, "environment");
      const conc = sc.filter((x) => x.rag === "Red" || x.rag === "Amber").sort((x, y) => RANK[y.rag] - RANK[x.rag]).slice(0, 3);
      return `<article class="sx-pcard" data-open="${esc(r.name)}">
        <div class="pc-top"><div><div class="pc-name">${esc(r.name)}</div><div class="pc-sub">${r.p.poap ? "POAP · " : ""}${r.p.raid ? "RAID · " : ""}${K.weeks(r.p.kpi).length} weeks of data</div></div>${R.chip(r.all.overall.rag)}</div>
        <div class="pc-bar"><span style="width:${(g / n) * 100}%;background:${C.green}"></span><span style="width:${(a / n) * 100}%;background:${C.amber}"></span><span style="width:${(rd / n) * 100}%;background:${C.red}"></span></div>
        <div class="pc-nums"><div><b>${g}</b><span>Green</span></div><div><b>${a}</b><span>Amber</span></div><div><b>${rd}</b><span>Red</span></div>
          <div><b>${ex.empty ? "—" : N.pct(ex.detail.execPct)}</b><span>Executed</span></div><div><b>${dd.empty || !dd.detail.open ? 0 : dd.detail.open.length}</b><span>Open defects</span></div><div><b>${en.empty ? "—" : N.pct1(en.detail.avail)}</b><span>Availability</span></div></div>
        <ul class="pc-conc">${conc.length ? conc.map((x) => `<li>${R.dot(x.rag)}<b>${esc(x.name)}</b> — ${esc(x.position)}</li>`).join("") : "<li>All reported KPIs are Green.</li>"}</ul>
      </article>`;
    }).join("")}</div>`;
  }

  function compare(rows) {
    const opts = K.KPI_META.map((m) => `<option value="${m.key}" ${m.key === st.kpi ? "selected" : ""}>${esc(m.name)}</option>`).join("");
    const have = rows.filter((r) => kpi(r, st.kpi).metric && kpi(r, st.kpi).metric.value != null);
    const meta = K.metaByKey[st.kpi];
    const tbl = U.table(["Portfolio", "Position", "RAG", "Change"], rows.slice().sort((a, b) => RANK[kpi(b, st.kpi).rag] - RANK[kpi(a, st.kpi).rag] || a.name.localeCompare(b.name)).map((r) => { const x = kpi(r, st.kpi); return [`<a data-open="${esc(r.name)}" data-page="${st.kpi}" style="cursor:pointer;font-weight:700">${esc(r.name)}</a>`, esc(x.position), R.chip(x.rag), U.deltaChip(x) || "—"]; }), { wrap: true });
    return `<div class="sx-toolbar"><label>KPI</label><select id="pg-kpi">${opts}</select><span class="sx-dim">${have.length} of ${rows.length} portfolios have data for ${esc(meta.name)}</span></div>
      <div class="sx-chartbox" style="height:${Math.max(240, have.length * 34 + 60)}px"><canvas id="pg-cmp"></canvas></div>${tbl}`;
  }

  function registers(rows) {
    const q = st.q.toLowerCase();
    const defs = {
      tsr: { label: "TSRs", heads: ["Portfolio", "TSR", "Received", "Returned", "Working days", "SLA"], get: (r) => (kpi(r, "tsr").detail.items || []).map((x) => [r.name, `<b>${esc(x.t.ref)}</b>`, D.fmt(x.t.received, false), x.open ? "Open" : D.fmt(x.t.returned, false), `<b>${x.wd}</b>`, x.over ? R.chip("Red", "Over") : R.chip("Green", "Within")]) },
      defects: { label: "Open defects", heads: ["Portfolio", "ID", "Summary", "Priority", "Status", "Age (WD)", "Owner"], get: (r) => { const a = kpi(r, "aging"); return a.empty ? [] : a.detail.open.map((x) => [r.name, `<b>${esc(x.d.id)}</b>`, esc(x.d.summary), esc(x.d.priority), esc(x.d.status), `<b style="color:${x.days > 10 ? C.red : "inherit"}">${x.days}</b>`, esc(x.d.owner)]); } },
      milestones: { label: "Milestones at risk", heads: ["Portfolio", "Milestone", "Due", "Status", "Delay (days)"], get: (r) => { const m = kpi(r, "milestones"); return m.empty ? [] : m.detail.ev.filter((e) => e.status === "Overdue" || e.rag === "Amber" || e.status === "Completed Late").map((e) => [r.name, esc(e.m.name), D.fmt(e.m.due, false), `${R.chip(e.rag, e.status)}`, e.delay || "—"]); } },
      raid: { label: "High RAID items", heads: ["Portfolio", "ID", "Item", "Rating", "Owner"], get: (r) => { const x = kpi(r, "raid"); if (x.empty) return []; const hi = (i) => i.rating === "High" || i.rating === "Very High"; return x.detail.risksOpen.filter(hi).map((i) => [r.name, `<b>${esc(i.id)}</b>`, esc(i.title), ATS.raid.ratingChip(i.rating), esc(i.owner)]).concat(x.detail.issuesOpen.filter(hi).map((i) => [r.name, `<b>${esc(i.id)}</b>`, esc(i.title), ATS.raid.ratingChip(i.rating), esc(i.owner)])); } },
    };
    const d = defs[st.reg];
    let data = [].concat(...rows.map(d.get)).filter((r) => !q || r.join(" ").toLowerCase().includes(q));
    const total = data.length; data = data.slice(0, 250);
    return `<div class="sx-toolbar"><div class="toggle-pill" id="pg-reg">${Object.keys(defs).map((k) => `<button data-r="${k}" class="${k === st.reg ? "active" : ""}">${defs[k].label}</button>`).join("")}</div>
      <div class="search-box"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg><input id="pg-q" type="text" placeholder="Filter all portfolios…" value="${esc(st.q)}"></div><span class="sx-dim">${total} row${total === 1 ? "" : "s"}</span></div>
      ${U.table(d.heads, data, { wrap: true })}`;
  }

  // ---------------------------------------------------------------- render
  P.render = function () {
    if (!root) return;
    ATS.destroyCharts("pg-");
    const ps = portfolios();
    if (!ps.length) { root.innerHTML = ATS.emptyCard("Programme Overview"); ATS.wireEmptyCard(root); return; }
    const per = periods(ps);
    let list = per[st.mode];
    if (!list.length) { st.mode = st.mode === "month" ? "week" : "month"; list = per[st.mode]; }
    if (!st.key || !list.includes(st.key)) {
      st.key = K.defaultKey(list, st.mode, ps.map((p) => p.kpi.snapshots));
    }
    const { rows, period } = compute(ps, st.mode, st.key);
    const a = aggregate(rows);
    last = { rows, period, mode: st.mode, agg: a };
    const idx = list.indexOf(st.key);
    const label = (k) => (st.mode === "month" ? D.fmtMonth(k) : "W/E " + D.fmt(k));
    const prog = (ps[0].kpi.config.programme || "").trim();
    root.innerHTML = `
      <div class="sx-head">
        <div><div class="sx-eyebrow">Programme overview</div><h2>${esc(prog || "All portfolios")} — ${esc(label(st.key))}</h2>
          <div class="sx-sub">${a.n} portfolio${a.n === 1 ? "" : "s"} · each fills its own weekly workbook — this page is derived from all of them</div></div>
        <div class="sx-controls">
          <div class="toggle-pill" id="pg-mode"><button data-m="month" class="${st.mode === "month" ? "active" : ""}">Monthly</button><button data-m="week" class="${st.mode === "week" ? "active" : ""}">Weekly</button></div>
          <button class="btn btn-outline btn-sm" id="pg-prev" ${idx <= 0 ? "disabled" : ""}>‹</button>
          <select id="pg-period">${list.slice().reverse().map((k) => `<option value="${k}" ${k === st.key ? "selected" : ""}>${esc(label(k))}</option>`).join("")}</select>
          <button class="btn btn-outline btn-sm" id="pg-next" ${idx >= list.length - 1 ? "disabled" : ""}>›</button>
          <button class="btn btn-primary btn-sm" id="pg-pptx">Export PowerPoint</button>
        </div>
      </div>
      <div class="sx-stats" style="margin-bottom:16px">
        ${U.stat(`<span style="color:${C.green}">${a.rag.Green}</span> / <span style="color:${C.amber}">${a.rag.Amber}</span> / <span style="color:${C.red}">${a.rag.Red}</span>`, "Green / Amber / Red portfolios")}
        ${U.stat(a.planned ? N.pct(a.exec / a.planned) : "—", "Test cases executed")}${U.stat(a.openDef, "Open defects")}${U.stat(a.tsrOver, "TSRs over SLA")}${U.stat(a.msOver, "Overdue milestones")}${U.stat(a.budget ? N.gbpK(a.spend) + " / " + N.gbpK(a.budget) : "—", "Spend / budget")}${U.stat(a.raidHigh, "High RAID items")}
      </div>
      <div class="toggle-pill" id="pg-tabs" style="margin-bottom:14px">${[["heatmap", "Heat-map"], ["cards", "Portfolio cards"], ["compare", "Compare a KPI"], ["registers", "Registers"]].map(([k, l]) => `<button data-t="${k}" class="${k === st.tab ? "active" : ""}">${l}</button>`).join("")}</div>
      <article class="sx-slide" id="pg-slide"><div id="pg-body">${st.tab === "heatmap" ? heatmap(rows) : st.tab === "cards" ? cards(rows) : st.tab === "compare" ? compare(rows) : registers(rows)}</div>
        <footer class="sx-slide-foot"><span>${esc(label(st.key))} · ${a.n} portfolios${ATS.isSample() ? ' · <b style="color:#b54708">SAMPLE DATA</b>' : ""}</span><span class="sx-foot-btns" data-export="pg-slide"><button class="btn btn-outline btn-sm" data-act="copy">Copy as picture</button><button class="btn btn-outline btn-sm" data-act="png">Download PNG</button></span></footer></article>`;

    // wiring
    const rerender = () => P.render();
    root.querySelectorAll("#pg-mode button").forEach((b) => (b.onclick = () => { st.mode = b.dataset.m; st.key = null; rerender(); }));
    root.querySelector("#pg-period").onchange = (e) => { st.key = e.target.value; rerender(); };
    root.querySelector("#pg-prev").onclick = () => { st.key = list[idx - 1]; rerender(); };
    root.querySelector("#pg-next").onclick = () => { st.key = list[idx + 1]; rerender(); };
    root.querySelector("#pg-pptx").onclick = () => (ATS.exportProgrammePptx ? ATS.exportProgrammePptx(last) : ATS.toast("Export not available", "err"));
    root.querySelectorAll("#pg-tabs button").forEach((b) => (b.onclick = () => { st.tab = b.dataset.t; rerender(); }));
    root.querySelectorAll("[data-open]").forEach((el) => (el.onclick = (e) => { e.stopPropagation(); open(el.dataset.open, el.dataset.page); }));
    const kp = root.querySelector("#pg-kpi"); if (kp) kp.onchange = (e) => { st.kpi = e.target.value; rerender(); };
    root.querySelectorAll("#pg-reg button").forEach((b) => (b.onclick = () => { st.reg = b.dataset.r; rerender(); }));
    const q = root.querySelector("#pg-q"); if (q) q.oninput = (e) => { st.q = e.target.value; const pos = e.target.selectionStart; rerender(); const n = root.querySelector("#pg-q"); n.focus(); n.setSelectionRange(pos, pos); };
    U.wireExports(root, "programme-" + st.mode + "-" + st.key);

    if (st.tab === "compare") {
      const meta = K.metaByKey[st.kpi];
      let have = rows.filter((r) => kpi(r, st.kpi).metric && kpi(r, st.kpi).metric.value != null);
      const m0 = have.length ? kpi(have[0], st.kpi).metric : null;
      const asc = m0 && m0.better === "down";
      have = have.sort((x, y) => (asc ? kpi(x, st.kpi).metric.value - kpi(y, st.kpi).metric.value : kpi(y, st.kpi).metric.value - kpi(x, st.kpi).metric.value));
      const f = m0 ? m0.fmt : "int";
      const val = (x) => { const v = kpi(x, st.kpi).metric.value; return f === "pct" || f === "pct1" ? Math.round(v * 1000) / 10 : f === "gbp" ? v : Math.round(v * 100) / 100; };
      ATS.chart("pg-cmp", root.querySelector("#pg-cmp"), { type: "bar", data: { labels: have.map((x) => x.name), datasets: [{ label: m0 ? m0.label : meta.name, data: have.map(val), backgroundColor: have.map((x) => R.color(kpi(x, st.kpi).rag)), borderRadius: 6, barThickness: 20 }] },
        options: Object.assign(U.base(), { indexAxis: "y", plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => kpi(have[c.dataIndex], st.kpi).position } } }, scales: { x: { grid: { color: C.grid }, beginAtZero: true, ticks: { callback: (v) => (f === "pct" || f === "pct1" ? v + "%" : v) } }, y: { grid: { display: false } } }, onClick: (e, els) => { if (els.length) open(have[els[0].index].name, st.kpi); } }) });
    }
  };
})();
