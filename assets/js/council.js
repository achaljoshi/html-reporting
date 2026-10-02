/* ============================================================
   Section framework for Monthly Council & Weekly Report
   ATS.makeSection(cfg) -> { mount(rootEl), render(), go(pageId), isActive }
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date, N = ATS.num, R = ATS.rag, K = ATS.kpi, U = ATS.ui, esc = ATS.dom.esc;

  // ------------------------------------------------------------------ shared pages
  ATS.pages = {};

  // KPI slide page
  ATS.pages.kpiSlide = function (env, key, eyebrow, opts) {
    opts = opts || {};
    const res = env.all.byKey[key];
    const fn = ATS.kpiPages[key];
    const body = fn(res, env);
    const quarter = res.cadence === "Quarterly" && env.period.type === "month" ? ` · quarter to date (${D.quarterOf(env.period.end).label})` : "";
    const sub = `${esc(env.period.label)}${quarter} · <b>${esc(res.position)}</b>${res.positionOverridden ? " (as written in Commentary)" : ""}`;
    const foot = `Source: ${esc(env.data.config.programName)} · data as of ${D.fmt(env.all.asOf)}${ATS.isSample() ? ' · <b style="color:#b54708">SAMPLE DATA</b>' : ""}`;
    const html = U.slide({ id: env.prefix + "slide-" + key, eyebrow: eyebrow || res.group, title: opts.title || res.name, sub, badge: U.ragBadge(res), left: body.left, right: U.narrative(res), foot });
    return { html, draw: body.draw };
  };

  // Overall scorecard
  ATS.pages.scorecard = function (env, o) {
    o = o || {};
    const all = env.all, sc = all.scorecard, prev = all.prev;
    const cnt = { Green: 0, Amber: 0, Red: 0, Grey: 0 };
    sc.forEach((x) => cnt[x.rag]++);
    const scored = cnt.Green + cnt.Amber + cnt.Red;
    const ov = all.overall;
    const hero = `
      <div class="sx-hero">
        <div class="ring">${U.ring(scored ? cnt.Green / scored : 0, "#22c55e", 96, scored ? Math.round((cnt.Green / scored) * 100) + "%" : "–")}</div>
        <div>
          <div class="sx-eyebrow" style="color:#7fe3de">${esc(o.eyebrow || "Programme status")}</div>
          <h3>${esc(env.data.config.programName)} — Overall <span class="chip ${R.cls(ov.rag)}" style="margin-left:8px;vertical-align:middle">${esc(ov.rag === "Grey" ? "N/A" : ov.rag)}</span></h3>
          <p class="clamp" title="Click to expand / collapse" onclick="this.classList.toggle('open')">${esc(ov.summary || autoSummary(all))}</p>
          ${ov.declared && ov.computed !== ov.declared ? `<p style="margin-top:6px;font-size:11.5px;color:#9fb4e8">Declared RAG <b>${esc(ov.declared)}</b> (your judgement) · computed from KPIs <b>${esc(ov.computed)}</b></p>` : ""}
        </div>
        <div class="counts"><div><b style="color:#4ade80">${cnt.Green}</b><span>Green</span></div><div><b style="color:#fbbf24">${cnt.Amber}</b><span>Amber</span></div><div><b style="color:#f87171">${cnt.Red}</b><span>Red</span></div><div><b style="color:#98a2b3">${cnt.Grey}</b><span>N/A</span></div></div>
      </div>`;
    // what changed
    let changes = "";
    if (prev) {
      const rank = { Grey: 0, Green: 1, Amber: 2, Red: 3 };
      const ch = sc.filter((x) => x.prev && x.prev.rag !== x.rag && x.prev.rag !== "Grey" && x.rag !== "Grey")
        .map((x) => `<span class="sx-change ${rank[x.rag] < rank[x.prev.rag] ? "up" : "down"}">${rank[x.rag] < rank[x.prev.rag] ? "▲" : "▼"} <b>${esc(x.name)}</b>: ${x.prev.rag} → ${x.rag}</span>`);
      changes = `<div class="sx-ctitle" style="margin-top:2px">Movement vs ${esc(all.prevPeriod.short)}</div><div class="sx-changes">${ch.length ? ch.join("") : '<span class="sx-change">No RAG changes</span>'}</div>`;
    }
    const mode = o.mode || "tiles";
    let series = null;
    if (mode === "tiles") series = K.weeklySeries(env.data, env.period.end, 8, { raid: env.raid, today: all.asOf > env.period.end ? env.period.end : undefined });
    const tiles = mode === "tiles"
      ? `<div class="sx-tiles">${sc.map((x) => U.tile(x, { vals: series.map((s) => (s.res.byKey[x.key].metric ? s.res.byKey[x.key].metric.value : null)) })).join("")}</div>`
      : U.table(["KPI", "Current position", "RAG", "Commentary"], sc.map((x) => [`<b>${esc(x.name)}</b>`, esc(x.position), R.chip(x.rag), esc(x.exec)]), { wrap: true });
    const toggle = `<div class="toggle-pill" id="${env.prefix}sc-mode"><button data-m="tiles" class="${mode === "tiles" ? "active" : ""}">Tiles</button><button data-m="table" class="${mode === "table" ? "active" : ""}">Table</button></div>`;
    const html = `<article class="sx-slide" id="${env.prefix}slide-scorecard">
      <header class="sx-slide-head"><div><div class="sx-slide-eyebrow">${esc(o.eyebrow || "Overview")}</div><h3>${esc(o.title || "Overall Dashboard")}</h3><div class="sx-slide-sub">${esc(env.period.label)} · click any tile to open the KPI page</div></div>${toggle}</header>
      ${hero}${changes}${tiles}
      <footer class="sx-slide-foot"><span>Source: ${esc(env.data.config.programName)} · data as of ${D.fmt(all.asOf)}${ATS.isSample() ? ' · <b style="color:#b54708">SAMPLE DATA</b>' : ""}</span>
        <span class="sx-foot-btns" data-export="${env.prefix}slide-scorecard"><button class="btn btn-outline btn-sm" data-act="copy">Copy as picture</button><button class="btn btn-outline btn-sm" data-act="png">Download PNG</button></span></footer></article>`;
    return {
      html,
      draw(root) {
        root.querySelectorAll("#" + env.prefix + "sc-mode button").forEach((b) => (b.onclick = () => { env.state.scMode = b.dataset.m; env.rerender(); }));
        root.querySelectorAll(".sx-tile").forEach((t) => (t.onclick = () => env.open(t.dataset.kpi)));
      },
    };
  };

  function autoSummary(all) {
    const sc = all.scorecard, red = sc.filter((x) => x.rag === "Red"), amb = sc.filter((x) => x.rag === "Amber");
    const parts = [];
    if (red.length) parts.push(`${red.length} KPI${red.length > 1 ? "s are" : " is"} Red (${red.map((x) => x.name).join(", ")})`);
    if (amb.length) parts.push(`${amb.length} Amber (${amb.map((x) => x.name).join(", ")})`);
    return parts.length ? parts.join("; ") + "." : "All reported KPIs are Green.";
  }


  // RAG history heat-map: every KPI x every month (or week)
  ATS.pages.ragHistory = function (env, o) {
    o = o || {};
    const isMonth = env.period.type === "month";
    const keys = (isMonth ? K.months(env.data) : K.weeks(env.data).slice(-12)).filter((k) => k <= (isMonth ? env.period.key : env.period.key));
    const cols = keys.map((k) => ({ k, res: K.computeAll(env.data, isMonth ? K.monthPeriod(k) : K.weekPeriod(k), { raid: env.raid, withPrev: false }) }));
    const scored = K.KPI_META.filter((m) => !m.extra || o.extra);
    const head = cols.map((c) => `<th>${esc(isMonth ? D.fmtMonth(c.k, false) : D.fmt(c.k, false))}</th>`).join("");
    const body = scored.map((m) => `<tr><td class="topic" style="cursor:pointer" data-kpi="${m.key}">${esc(m.name)}</td>${cols.map((c) => {
      const r = c.res.byKey[m.key];
      return `<td><span class="sx-cellrag" style="background:${R.color(r.rag)}" title="${esc(m.name + " — " + (isMonth ? D.fmtMonth(c.k) : "W/E " + D.fmt(c.k)) + ": " + r.position)}">${U.ragLetter(r.rag)}</span></td>`;
    }).join("")}</tr>`).join("");
    const overall = `<tr><td class="topic"><b>Overall (declared)</b></td>${cols.map((c) => `<td><span class="sx-cellrag" style="background:${R.color(c.res.overall.rag)}">${U.ragLetter(c.res.overall.rag)}</span></td>`).join("")}</tr>`;
    const left = `<div class="table-wrap"><table class="sx-matrix"><thead><tr><th style="text-align:left">KPI</th>${head}</tr></thead><tbody>${overall}${body}</tbody></table></div>
      <p class="sx-empty-note">Each dot is recalculated from your logs <i>as of the end of that ${isMonth ? "month" : "week"}</i> — hover for the position, click a KPI name to open its page.</p>`;
    const html = U.slide({ id: env.prefix + "slide-history", eyebrow: o.eyebrow || "Trend", title: o.title || "RAG History", sub: isMonth ? "Month by month" : "Last 12 weeks", left, single: true, foot: "Source: " + esc(env.data.config.programName) + (ATS.isSample() ? ' · <b style="color:#b54708">SAMPLE DATA</b>' : "") });
    return { html, draw(root) { root.querySelectorAll("[data-kpi]").forEach((c) => (c.onclick = () => env.open(c.dataset.kpi))); } };
  };

  ATS.pages.lessons = function (env) {
    const rows = env.data.lessons.slice().sort((a, b) => ((a.date || "") < (b.date || "") ? 1 : -1));
    const tbl = rows.length ? U.table(["Date", "Category", "Lesson learned", "Improvement action", "Owner", "Status"], rows.map((r) => [r.date ? D.fmt(r.date, false) : "—", esc(r.category), esc(r.lesson), esc(r.action), esc(r.owner), esc(r.status)]), { wrap: true })
      : `<div class="sx-callout info">No lessons recorded yet. Add rows to the <b>Lessons_Learned</b> sheet of your workbook.</div>`;
    const html = U.slide({ id: env.prefix + "slide-lessons", eyebrow: "Risks, Issues & Lessons — Lessons Learned", title: "Lessons Learned", sub: "Post-transition lessons and improvement actions", left: tbl, single: true, foot: "Source: Lessons_Learned sheet" + (ATS.isSample() ? ' · <b style="color:#b54708">SAMPLE DATA</b>' : "") });
    return { html, draw() {} };
  };

  ATS.pages.checks = function (env) {
    const q = K.quality(env.data, env.all.asOf);
    const d = env.data;
    const meta = ATS.store.meta;
    const files = (meta.files || []).map((f) => `<span class="chip green">${esc(f.kind.toUpperCase())}</span> ${esc(f.name)} <span class="sx-dim">— ${esc(f.note || "")}</span>`).join("<br>");
    const left = `<div class="sx-stats">${U.stat(d.snapshots.length, "Weekly snapshots")}${U.stat(d.defects.length, "Defects")}${U.stat(d.tsrs.length, "TSRs")}${U.stat(d.milestones.length, "Milestones")}${U.stat(d.readiness.length, "Readiness rows")}</div>
      <div class="sx-panel"><h5>Data checks <span class="tag">${q.length ? q.length + " to review" : "all clear"}</span></h5>${q.length ? `<ul class="sx-quality">${q.map((x) => `<li><b>${esc(x.sheet)}</b> — ${esc(x.msg)}</li>`).join("")}</ul>` : '<p>No problems found in the workbook. Dates, statuses and counts are consistent.</p>'}</div>
      <div class="sx-panel"><h5>Loaded files</h5><p>${files || "Embedded sample data"}</p><p style="margin-top:8px" class="sx-dim">Loaded ${meta.loadedAt ? new Date(meta.loadedAt).toLocaleString() : "—"} · Source: ${esc(meta.source)}</p></div>`;
    return { html: U.slide({ id: env.prefix + "slide-checks", eyebrow: "Appendix", title: "Data Quality & Provenance", sub: "Why you can trust the numbers", left, single: true, foot: "Checks run every time data is loaded" }), draw() {} };
  };

  // ------------------------------------------------------------------ factory
  ATS.makeSection = function (cfg) {
    const state = { key: null, page: null, scMode: "tiles", compare: true };
    let root = null;
    const sec = { state };

    function periodList(data) {
      return cfg.kind === "month" ? K.months(data) : K.weeks(data);
    }
    function periodOf(key) { return cfg.kind === "month" ? K.monthPeriod(key) : K.weekPeriod(key); }
    function labelOf(key) { return cfg.kind === "month" ? D.fmtMonth(key) : "W/E " + D.fmt(key); }

    sec.mount = function (el) { root = el; };
    sec.root = () => root;

    sec.render = function () {
      if (!root) return;
      const store = ATS.store, data = store.kpi;
      ATS.destroyCharts(cfg.id + "-");
      if (data && !data.snapshots.length) {
        root.innerHTML = `<div class="card card-pad sx-load-card"><div class="sx-eyebrow">${esc(cfg.eyebrow)}</div><h2 style="margin:6px 0 8px">Your workbook is loaded but has no weekly rows yet</h2>
          <p class="text-sub" style="line-height:1.6;margin:0">Add the first week to the <b>Weekly_Snapshot</b> sheet of <b>${esc(data.config.programName)}</b> (one row per week ending Friday), save, then click <b>Refresh data</b>. Rows marked <i>EXAMPLE</i> in the Notes column are ignored — delete them once you have entered real data.</p>
          <div class="row"><button class="btn btn-primary" data-load="refresh">Refresh data</button><button class="btn btn-outline" data-load="folder">Load a different folder</button></div></div>`;
        ATS.wireEmptyCard(root); return;
      }
      if (!data) { root.innerHTML = ATS.emptyCard(cfg.eyebrow); ATS.wireEmptyCard(root); return; }
      const list = periodList(data);
      if (!list.length) { root.innerHTML = ATS.emptyCard(cfg.eyebrow); ATS.wireEmptyCard(root); return; }
      if (!state.key || !list.includes(state.key)) {
        state.key = list[list.length - 1];
        // monthly: don't default to a month that only has one or two weeks of data
        if (cfg.kind === "month" && list.length > 1) {
          const n = data.snapshots.filter((s) => s.weekEnding.slice(0, 7) === state.key).length;
          if (n < 3) state.key = list[list.length - 2];
        }
      }
      const period = periodOf(state.key);
      const all = K.computeAll(data, period, { raid: store.raid });
      const env = { data, raid: store.raid, poap: store.poap, period, all, prefix: cfg.id + "-", state, store,
        rerender: () => sec.render(), open: (page) => sec.go(page) };
      env.ctx = env;
      const pages = cfg.pages(env);
      if (!state.page || !pages.some((p) => p.id === state.page)) state.page = pages[0].id;

      // head
      const idx = list.indexOf(state.key);
      root.innerHTML = `
        <div class="sx-head">
          <div><div class="sx-eyebrow">${esc(cfg.eyebrow)}</div><h2>${esc(cfg.kind === "month" ? D.fmtMonth(state.key) : "Week ending " + D.fmt(state.key))}</h2>
            <div class="sx-sub">${esc(data.config.programName)}${data.config.project ? " · " + esc(data.config.project) : ""} · ${esc(cfg.tagline)}</div></div>
          <div class="sx-controls">
            <button class="btn btn-outline btn-sm" id="${env.prefix}prev" ${idx <= 0 ? "disabled" : ""} title="Previous">‹</button>
            <select id="${env.prefix}period">${list.slice().reverse().map((k) => `<option value="${k}" ${k === state.key ? "selected" : ""}>${esc(labelOf(k))}</option>`).join("")}</select>
            <button class="btn btn-outline btn-sm" id="${env.prefix}next" ${idx >= list.length - 1 ? "disabled" : ""} title="Next">›</button>
            <button class="btn btn-primary btn-sm" id="${env.prefix}pptx" title="Build a PowerPoint from this data — no screenshots needed">
              <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12M7 10l5 5 5-5M4 20h16"/></svg> Export PowerPoint</button>
          </div>
        </div>
        <div class="sx-layout"><nav class="sx-rail" id="${env.prefix}rail"></nav><div class="sx-stage" id="${env.prefix}stage"></div></div>`;

      // rail
      const rail = root.querySelector("#" + env.prefix + "rail");
      let g = "", rh = "";
      pages.forEach((p) => {
        if (p.group !== g) { g = p.group; rh += `<h6>${esc(g)}</h6>`; }
        const rg = p.kpi ? env.all.byKey[p.kpi].rag : p.ragOf ? p.ragOf(env) : null;
        rh += `<button data-page="${p.id}" class="${p.id === state.page ? "active" : ""}">${rg ? R.dot(rg) : ""}${esc(p.label)}${p.n ? `<span class="sx-n">${esc(p.n)}</span>` : ""}</button>`;
      });
      rail.innerHTML = rh;
      rail.querySelectorAll("button").forEach((b) => (b.onclick = () => sec.go(b.dataset.page)));

      // controls
      root.querySelector("#" + env.prefix + "period").onchange = (e) => { state.key = e.target.value; sec.render(); };
      root.querySelector("#" + env.prefix + "prev").onclick = () => { state.key = list[idx - 1]; sec.render(); };
      root.querySelector("#" + env.prefix + "next").onclick = () => { state.key = list[idx + 1]; sec.render(); };
      root.querySelector("#" + env.prefix + "pptx").onclick = () => (ATS.exportPptx ? ATS.exportPptx(cfg.kind, env, pages) : ATS.toast("PowerPoint export not available", "err"));

      // stage
      const stage = root.querySelector("#" + env.prefix + "stage");
      const page = pages.find((p) => p.id === state.page);
      let out;
      try { out = page.render(env); } catch (e) { console.error(e); out = { html: `<div class="sx-callout red">This page failed to render: ${esc(e.message)}</div>`, draw() {} }; }
      stage.innerHTML = Array.isArray(out) ? out.map((o) => o.html).join('<div style="height:18px"></div>') : out.html;
      (Array.isArray(out) ? out : [out]).forEach((o) => o.draw && o.draw(stage));
      U.wireExports(stage, cfg.id + "-" + state.page + "-" + state.key);
      env.pages = pages;
      sec._env = env;
    };

    sec.go = function (pageId) {
      state.page = pageId;
      sec.render();
      if (root) window.scrollTo({ top: 0, behavior: "auto" });
    };

    sec.step = function (dir) {
      const env = sec._env; if (!env) return;
      const i = env.pages.findIndex((p) => p.id === state.page), j = i + dir;
      if (j >= 0 && j < env.pages.length) sec.go(env.pages[j].id);
    };
    return sec;
  };

  // ------------------------------------------------------------------ empty / load card
  ATS.emptyCard = function (title) {
    return `<div class="card card-pad sx-load-card">
      <div class="sx-eyebrow">${ATS.dom.esc(title || "ATS")}</div><h2 style="margin:6px 0 8px">Load your program data</h2>
      <p class="text-sub" style="line-height:1.6;margin:0">Pick the folder that holds your <b>ATS_Weekly_Data.xlsx</b>, <b>ATS_POAP.xlsx</b> and <b>ATS RAID Log.xlsx</b>. Everything is read locally on this machine — nothing is uploaded.</p>
      <div class="row"><button class="btn btn-primary" data-load="folder">Load program folder</button><button class="btn btn-outline" data-load="files">Choose files…</button>${window.ATS_SAMPLE ? '<button class="btn btn-outline" data-load="sample">Use sample data</button>' : ""}</div></div>`;
  };
  ATS.wireEmptyCard = function (rootEl) {
    rootEl.querySelectorAll("[data-load]").forEach((b) => (b.onclick = () => ATS.shell.load(b.dataset.load)));
  };
})();
