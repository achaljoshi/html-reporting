/* ============================================================
   Weekly Report — same data as the Monthly Council, sliced by week
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date, N = ATS.num, R = ATS.rag, K = ATS.kpi, U = ATS.ui, esc = ATS.dom.esc;
  const C = U.C;

  const kpiPage = (key, eyebrow, label, title) => ({ id: key, label, kpi: key, render: (env) => ATS.pages.kpiSlide(env, key, eyebrow, { title }) });
  const inRange = (d, a, b) => d && d >= a && d <= b;

  // ---------------------------------------------------------------- what changed
  function changes(env) {
    const { data, all, period } = env, good = [], watch = [];
    const ex = all.byKey.execution, exp = ex.prev;
    if (!ex.empty) {
      const t = ex.detail.total;
      const dx = t.executed - ex.detail.prevTotal.executed;
      if (dx > 0) good.push(`${N.int(dx)} more test cases executed (${N.int(t.executed)} of ${N.int(t.planned)} — ${N.pct(ex.detail.execPct)}); pass rate ${N.pct(ex.detail.passRate)}.`);
      else if (t.planned) watch.push(`No additional test cases were executed this week (${N.int(t.executed)} of ${N.int(t.planned)}).`);
      const db = t.blocked - ex.detail.prevTotal.blocked;
      if (db < 0) good.push(`Blocked test cases reduced by ${N.int(-db)} to ${N.int(t.blocked)}.`);
      else if (db > 0) watch.push(`Blocked test cases increased by ${N.int(db)} to ${N.int(t.blocked)}.`);
      else if (t.blocked) watch.push(`${N.int(t.blocked)} test cases remain blocked.`);
    }
    const raised = data.defects.filter((d) => inRange(d.raised, period.start, period.end));
    const resolved = data.defects.filter((d) => inRange(d.resolved, period.start, period.end) && d.status !== "Rejected");
    const open = data.defects.filter((d) => d.raised <= all.asOf && d.status !== "Rejected" && !(d.resolved && d.resolved <= all.asOf));
    if (resolved.length) good.push(`${N.plural(resolved.length, "defect")} resolved: ${resolved.map((d) => d.id).join(", ")}.`);
    if (raised.length) watch.push(`${N.plural(raised.length, "new defect")} raised: ${raised.map((d) => `${d.id} (${d.priority.split(" ")[0]})`).join(", ")}.`);
    if (open.length) watch.push(`${open.length} defect${open.length > 1 ? "s" : ""} open (${open.filter((d) => d.priority === "P1 High").length} High) — oldest: ${open.slice().sort((a, b) => (a.raised < b.raised ? -1 : 1))[0].id}.`);
    else good.push("No open defects.");
    const ms = all.byKey.milestones;
    if (!ms.empty) {
      const done = ms.detail.ev.filter((e) => inRange(e.m.completed, period.start, period.end));
      if (done.length) good.push(`Milestone${done.length > 1 ? "s" : ""} completed: ${done.map((e) => e.m.name).join(", ")}.`);
      ms.detail.ev.filter((e) => e.status === "Overdue").forEach((e) => watch.push(`${e.m.name} is ${e.delay} days overdue${e.m.forecast ? ` (forecast ${D.fmt(e.m.forecast, false)})` : ""}.`));
    }
    const env2 = all.byKey.environment;
    if (!env2.empty) (env2.rag === "Green" ? good : watch).push(`Environment downtime ${N.round1(env2.detail.down)}h this week (${N.pct1(env2.detail.avail)} availability).`);
    all.scorecard.forEach((x) => {
      if (!x.prev || x.prev.rag === x.rag || x.rag === "Grey" || x.prev.rag === "Grey") return;
      const rank = { Green: 1, Amber: 2, Red: 3 };
      (rank[x.rag] < rank[x.prev.rag] ? good : watch).push(`${x.name} moved ${x.prev.rag} → ${x.rag}.`);
    });
    if (all.byKey.raid && !all.byKey.raid.empty && all.byKey.raid.detail.newInPeriod) watch.push(`${all.byKey.raid.detail.newInPeriod} new RAID item(s) logged this week.`);
    return { good, watch };
  }

  // ---------------------------------------------------------------- look-ahead (POAP + milestones)
  ATS.lookahead = function (env, days) {
    days = days || 28;
    const from = D.add(env.all.asOf, 1), to = D.add(env.all.asOf, days), rows = [];
    env.data.milestones.filter((m) => !m.completed && inRange(m.due, from, to)).forEach((m) => rows.push({ date: m.due, kind: "Milestone", item: m.name, where: m.topic || "Programme", rag: D.diff(env.all.asOf, m.due) <= env.data.th.msWindow ? "Amber" : "Green" }));
    const p = env.poap;
    if (p) {
      p.plan.forEach((b) => {
        if (b.start && inRange(b.start, from, to)) rows.push({ date: b.start, kind: "Starts", item: `${b.type}: ${b.item}`, where: `${b.pillar} › ${b.topic}`, rag: b.rag || "Green" });
        if (b.end && inRange(b.end, from, to) && b.type !== "Holiday / Freeze") rows.push({ date: b.end, kind: "Ends", item: `${b.type}: ${b.item}`, where: `${b.pillar} › ${b.topic}`, rag: b.rag || "Green" });
      });
      (p.milestones || []).filter((m) => inRange(m.date, from, to)).forEach((m) => rows.push({ date: m.date, kind: "Milestone", item: m.name, where: `${m.pillar || ""} › ${m.topic || ""}`, rag: m.status === "At Risk" || m.status === "Missed" ? "Red" : "Green" }));
    }
    return rows.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  };

  function lookSlide(env) {
    const la = ATS.lookahead(env, 28).slice(0, 12);
    const pill = (kind) => `<span class="chip ${kind === "Milestone" ? "amber" : kind === "Starts" ? "green" : "grey"}">${kind}</span>`;
    return `<article class="sx-slide" id="${env.prefix}slide-look"><header class="sx-slide-head"><div><div class="sx-slide-eyebrow">Next 4 weeks</div><h3>Look-ahead</h3><div class="sx-slide-sub">${D.fmt(D.add(env.all.asOf, 1), false)} – ${D.fmt(D.add(env.all.asOf, 28))}${env.poap ? " · from POAP + milestones" : " · load a POAP workbook for more"}</div></div></header>
          ${U.table(["Date", "What", "Where"], la.map((r) => [D.fmt(r.date, false), `${pill(r.kind)} ${esc(r.item)}`, esc(r.where)]), { wrap: true })}
          <footer class="sx-slide-foot"><span>${la.length ? "" : "Nothing scheduled in the next 4 weeks."}</span><span class="sx-foot-btns" data-export="${env.prefix}slide-look"><button class="btn btn-outline btn-sm" data-act="copy">Copy as picture</button><button class="btn btn-outline btn-sm" data-act="png">Download PNG</button></span></footer></article>`;
  }
  function changesSlide(env) {
    const ch = changes(env);
    return `<article class="sx-slide" id="${env.prefix}slide-changes"><header class="sx-slide-head"><div><div class="sx-slide-eyebrow">Since last week</div><h3>What changed</h3></div></header>
          <div class="sx-panel hi" style="margin-bottom:10px"><h5>Progress <span class="tag">auto</span></h5>${U.list(ch.good)}</div>
          <div class="sx-panel co"><h5>Watch-outs <span class="tag">auto</span></h5>${U.list(ch.watch)}</div>
          <footer class="sx-slide-foot"><span>Generated from this week's data</span><span class="sx-foot-btns" data-export="${env.prefix}slide-changes"><button class="btn btn-outline btn-sm" data-act="copy">Copy as picture</button><button class="btn btn-outline btn-sm" data-act="png">Download PNG</button></span></footer></article>`;
  }
  function glancePage(env) {
    const head = ["execution", "dde", "aging", "coverage", "environment", "milestones", "raid", "readiness", "commercial", "resource"];
    const e2 = Object.assign({}, env, { all: Object.assign({}, env.all, { scorecard: head.map((k) => env.all.byKey[k]) }) });
    const sc = ATS.pages.scorecard(e2, { mode: "tiles", title: "Week at a Glance", eyebrow: "Weekly Report" });
    const html2 = `<div class="sx-two" style="margin-top:18px">${changesSlide(env)}${lookSlide(env)}</div>`;
    return [{ html: sc.html, draw: sc.draw }, { html: html2, draw() {} }];
  }

  // ---------------------------------------------------------------- trends
  function trendsPage(env) {
    const p = env.prefix;
    const series = K.weeklySeries(env.data, env.period.end, 12, { raid: env.raid });
    const lab = series.map((s) => D.fmt(s.we, false));
    const g = (key, fn) => series.map((s) => { const r = s.res.byKey[key]; return r && !r.empty ? fn(r) : null; });
    const openDefects = series.map((s) => env.data.defects.filter((d) => d.raised <= s.we && d.status !== "Rejected" && !(d.resolved && d.resolved <= s.we)).length);
    const cards = [
      ["tr-exec", "Test execution", "cumulative"], ["tr-cov", "Risk based coverage", "%"], ["tr-env", "Environment availability", "%"],
      ["tr-def", "Open defects", "count"], ["tr-res", "Resources", "FTE"], ["tr-spend", "Spend vs forecast", "£"],
    ];
    const left = `<div class="sx-grid3">${cards.map((c) => U.chartBox(p + c[0], 210, c[1], c[2])).join("")}</div>`;
    const html = U.slide({ id: p + "slide-trends", eyebrow: "Last 12 weeks", title: "Weekly Trends", sub: "Every number below comes from the weekly rows you have already entered — no extra work", left, single: true, foot: "Trends are computed per week-ending from Weekly_Snapshot, Topic_Progress and Defect_Log" });
    return {
      html,
      draw(root) {
        const cv = (id) => root.querySelector("#" + p + id);
        ATS.chart(p + "tr-exec", cv("tr-exec"), { type: "bar", data: { labels: lab, datasets: [
          { label: "Passed", data: g("execution", (r) => r.detail.total.passed), backgroundColor: C.teal, stack: "a" }, { label: "Failed", data: g("execution", (r) => r.detail.total.failed), backgroundColor: C.red, stack: "a" }, { label: "Blocked", data: g("execution", (r) => r.detail.total.blocked), backgroundColor: C.amber, stack: "a" } ] },
          options: Object.assign(U.base(), { scales: { x: Object.assign(U.axes().x, { stacked: true }), y: { stacked: true, grid: { color: C.grid } } } }) });
        const ln = (label, data, color, extra) => Object.assign({ label, data, borderColor: color, backgroundColor: color + "22", tension: 0.3, pointRadius: 2.5, borderWidth: 2, spanGaps: true }, extra || {});
        ATS.chart(p + "tr-cov", cv("tr-cov"), { type: "line", data: { labels: lab, datasets: [ln("High", g("coverage", (r) => Math.round(r.detail.levels.high.pct * 1000) / 10), C.red), ln("Medium", g("coverage", (r) => Math.round(r.detail.levels.med.pct * 1000) / 10), C.amber)] }, options: Object.assign(U.base(), { scales: { x: U.axes().x, y: { max: 100, beginAtZero: true, grid: { color: C.grid } } } }) });
        ATS.chart(p + "tr-env", cv("tr-env"), { type: "line", data: { labels: lab, datasets: [ln("Availability %", g("environment", (r) => Math.round(r.detail.avail * 1000) / 10), C.teal, { fill: true })] }, options: Object.assign(U.base(), { plugins: { legend: { display: false } }, scales: { x: U.axes().x, y: { min: 80, max: 100, grid: { color: C.grid } } } }) });
        ATS.chart(p + "tr-def", cv("tr-def"), { type: "bar", data: { labels: lab, datasets: [{ label: "Open defects", data: openDefects, backgroundColor: C.indigo, borderRadius: 4 }] }, options: Object.assign(U.base(), { plugins: { legend: { display: false } }, scales: Object.assign(U.axes(), { y: { grid: { color: C.grid }, beginAtZero: true, ticks: { precision: 0 } } }) }) });
        ATS.chart(p + "tr-res", cv("tr-res"), { type: "line", data: { labels: lab, datasets: [ln("Planned", g("resource", (r) => r.detail.plan), C.navy, { stepped: true, borderDash: [5, 4], backgroundColor: "transparent" }), ln("Actual", g("resource", (r) => r.detail.act), C.teal, { stepped: true })] }, options: Object.assign(U.base(), { scales: Object.assign(U.axes(), { y: { grid: { color: C.grid }, beginAtZero: true } }) }) });
        ATS.chart(p + "tr-spend", cv("tr-spend"), { type: "line", data: { labels: lab, datasets: [ln("Spend", g("commercial", (r) => r.detail.spend), C.teal, { fill: true }), ln("Forecast", g("commercial", (r) => r.detail.forecast), C.indigo, { borderDash: [5, 4], backgroundColor: "transparent" })] }, options: Object.assign(U.base(), { scales: { x: U.axes().x, y: { grid: { color: C.grid }, ticks: { callback: (v) => "£" + Math.round(v / 1000) + "k" } } } }) });
      },
    };
  }

  ATS.weekly = ATS.makeSection({
    id: "w", kind: "week", eyebrow: "Weekly Report", tagline: "Weekly delivery & quality report",
    pages(env) {
      const pg = [], add = (group, p) => { p.group = group; pg.push(p); };
      add("This week", { id: "glance", label: "Week at a Glance", render: glancePage });
      add("This week", { id: "trends", label: "Trends (12 weeks)", render: trendsPage });
      add("This week", { id: "history", label: "RAG History", render: (e) => ATS.pages.ragHistory(e, { eyebrow: "Week-by-week", title: "RAG History by Week", extra: true }) });
      add("Delivery", kpiPage("execution", "Delivery — Test Execution", "Test Execution"));
      add("Delivery", kpiPage("dde", "Delivery — Defects", "Defects", "Defects"));
      add("Delivery", kpiPage("aging", "Delivery — Defect Aging", "Defect Aging"));
      add("Delivery", kpiPage("milestones", "Delivery — Milestones", "Milestones"));
      add("Delivery", kpiPage("readiness", "Delivery — SIT Readiness", "SIT Readiness"));
      add("Quality", kpiPage("coverage", "Quality — Risk Based Coverage", "Risk Based Coverage"));
      add("Quality", kpiPage("environment", "Quality — Test Environment", "Test Environment"));
      add("Quality", kpiPage("automation", "Quality — Automation", "Automation"));
      add("Governance", kpiPage("raid", "Governance — RAID", "RAID Log"));
      add("Governance", kpiPage("tsr", "Governance — TSR Impact Assessment", "TSR Impact"));
      add("Commercial", kpiPage("commercial", "Commercials", "Commercials"));
      add("Commercial", kpiPage("resource", "Resourcing", "Resourcing"));
      add("Look ahead", { id: "look", label: "4-Week Look-ahead", render: (e) => ({ html: lookSlide(e), draw() {} }) });
      add("Appendix", { id: "checks", label: "Data Quality", render: (e) => ATS.pages.checks(e) });
      return pg;
    },
  });
})();
