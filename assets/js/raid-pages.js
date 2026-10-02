/* ============================================================
   RAID log pages — one page per RAID sheet, each with its own overview, matrix / heat-map and open items:
   Risks · Issues · Assumptions · Dependencies · Decisions   (Lessons Learned is its own page)
   ATS.raidPages(kind)  -> page render function (env) => { html, draw }
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date, N = ATS.num, U = ATS.ui, RA = ATS.raid, esc = ATS.dom.esc;
  const lc = (s) => N.str(s).toLowerCase();
  const trunc = (s, n) => { s = N.str(s); return s.length > n ? s.slice(0, n - 1) + "…" : s; };
  const LV = ["Very High", "High", "Medium", "Low", "Very Low"];
  const levelOf = (s) => LV.find((x) => x.toLowerCase() === lc(s)) || "Unrated";
  const LEVEL_COL = { "Very High": "#ef5b5b", High: "#ff9d5c", Medium: "#ffd66b", Low: "#c9e58f", "Very Low": "#9ad8b0", Unrated: "#d0d5dd" };
  const TEAL = "#0ea5a0";
  const dfmt = (d) => (d ? D.fmt(d, false) : "—");
  const chip = (txt, cls) => `<span class="chip ${cls || "grey"}">${esc(txt || "—")}</span>`;
  const statusChip = (s) => { const k = lc(s); return chip(s, /^(closed|resolved|approved|confirmed correct)/.test(k) ? "green" : /^(rejected|confirmed incorrect)/.test(k) ? "red" : /^(pending|unconfirmed|proposed|open|resolution)/.test(k) ? "amber" : "grey"); };
  const late = (iso, asOf) => iso && iso < asOf;
  const bars = (list) => RA.bandBars(list.map((x) => ({ label: x.label, color: x.color || TEAL, n: x.n })));
  const countBy = (arr, fn, order) => {
    const m = {}; arr.forEach((x) => { const k = fn(x) || "—"; m[k] = (m[k] || 0) + 1; });
    const keys = order ? order.filter((k) => m[k]).concat(Object.keys(m).filter((k) => !order.includes(k))) : Object.keys(m).sort((a, b) => m[b] - m[a]);
    return keys.map((k) => ({ label: k, n: m[k] }));
  };
  const pan = (hi, co, ac, ex) => U.narrative({ highlights: hi, concern: co, actions: ac, exec: ex, src: {} });
  const list = (arr, n) => arr.slice(0, n || 3).map((x) => x.id).join(", ") + (arr.length > (n || 3) ? "…" : "");

  // generic matrix (rows x cols), clickable cells
  function matrix(o) {
    let h = `<div class="sx-mx" id="${o.id}" style="grid-template-columns:${o.rowW || 92}px repeat(${o.cols.length},minmax(0,1fr))"><div class="sx-mx-corner">${esc(o.corner || "")}</div>`;
    h += o.cols.map((c) => `<div class="sx-mx-col">${esc(c.label)}</div>`).join("");
    o.rows.forEach((r, ri) => {
      h += `<div class="sx-mx-row">${esc(r.label)}</div>`;
      o.cols.forEach((c, ci) => {
        const x = o.cell(r, c), n = x.items.length;
        h += `<button type="button" class="sx-mx-cell${n ? " has" : ""}" data-r="${ri}" data-c="${ci}" style="background:${x.color}${n ? "" : "33"}" title="${esc(r.label + " × " + c.label)}${n ? " — " + esc(x.items.map((i) => i.id).join(", ")) : ""}">${n ? `<b>${n}</b>` : ""}</button>`;
      });
    });
    return h + "</div>";
  }

  // click a matrix cell -> show just those items in the table; button resets
  function wireFilter(root, mxId, boxId, itemsFor, render, title, all) {
    const mx = root.querySelector("#" + mxId), box = root.querySelector("#" + boxId);
    if (!mx || !box) return;
    mx.addEventListener("click", (e) => {
      const b = e.target.closest(".sx-mx-cell, .sx-heat-cell"); if (!b) return;
      const sel = itemsFor(b);
      box.innerHTML = `<div class="sx-ctitle">${esc(title(b, sel))} <button class="btn btn-outline btn-sm" data-reset style="margin-left:8px">Show all</button></div>${render(sel)}`;
      box.querySelector("[data-reset]").onclick = () => { box.innerHTML = all(); };
    });
  }

  function frame(env, kind, o) {
    const key = "raid-" + kind;
    const foot = `Source: RAID log · ${esc(env.data.config.programName)} · as of ${D.fmt(env.all.asOf)}${ATS.isSample() ? ' · <b style="color:#b54708">SAMPLE DATA</b>' : ""}`;
    const rag = o.rag || "Grey";
    return U.slide({ id: env.prefix + "slide-" + key, eyebrow: "Risks, Issues & Lessons — RAID log", title: o.title, sub: `${esc(env.period.label)} · <b>${esc(o.position)}</b>`, badge: U.ragBadge({ rag, autoRag: rag }), left: o.left, right: o.right, foot });
  }
  function missing(env, kind, title, msg) {
    return { html: U.slide({ id: env.prefix + "slide-raid-" + kind, eyebrow: "Risks, Issues & Lessons — RAID log", title, left: `<div class="sx-callout info">${msg}</div>`, single: true, foot: "" }), draw() {} };
  }
  const NO_RAID = "No RAID log is loaded for this portfolio. Put its <b>ATS_RAID_Log.xlsx</b> (copy of <code>templates/ATS_RAID_Log_Template.xlsx</code>) in the same folder as the weekly workbook and click <b>Refresh data</b>.";
  const rate = (o) => (o.veryHigh ? "Red" : o.high ? "Amber" : "Green");

  // ----------------------------------------------------------------------------------------------- RISKS
  function risks(env) {
    if (!env.raid) return missing(env, "risks", "Risks", NO_RAID);
    const asOf = env.all.asOf, p = env.prefix + "rk-";
    const all = env.raid.risks.filter((r) => !r.archived), open = all.filter(RA.isOpenRisk), closed = all.filter((r) => !RA.isOpenRisk(r));
    const mit = open.filter((r) => /mitigat/i.test(r.status)), vh = open.filter((r) => r.rating === "Very High"), hi = open.filter((r) => r.rating === "High");
    const overdue = open.filter((r) => late(r.review, asOf)), soon = open.filter((r) => r.review && r.review >= asOf && r.review <= D.add(asOf, 14));
    const tbl = (items) => U.table(["ID", "Risk", "Category", "L × I", "Rating", "Status", "Owner", "Review", "Mitigation"], items.map((r) => [`<b>${esc(r.id)}</b>`, esc(trunc(r.title, 110)), esc(r.category || "—"), r.l && r.i ? `${r.l}×${r.i}=${r.score}` : "—", RA.ratingChip(r.rating), statusChip(r.status), esc(trunc(r.owner, 28)), `<span style="${late(r.review, asOf) ? "color:#b42318;font-weight:700" : ""}">${dfmt(r.review)}</span>`, `<span class="sx-dim">${esc(trunc(r.mitigation, 160))}</span>`]), { wrap: true });
    const sorted = open.slice().sort((a, b) => (b.score || 0) - (a.score || 0));
    const allTbl = () => `<div class="sx-ctitle">Open risks <span class="sx-csub">· ${open.length}, highest score first (mitigated included)</span></div>${open.length ? tbl(sorted) : '<div class="sx-empty-note">No open risks.</div>'}`;
    const left = `<div class="sx-stats">${U.stat(open.length, "Open risks")}${U.stat(vh.length + hi.length, "High / Very High")}${U.stat(mit.length, "Mitigated")}${U.stat(closed.length, "Closed")}${U.stat(overdue.length, "Reviews overdue")}</div>
      <div class="sx-grid2"><div><div class="sx-ctitle">Risk heat-map <span class="sx-csub">· likelihood × impact, click a cell</span></div><div id="${p}heat">${RA.heatmap(open)}</div></div>
      <div><div class="sx-ctitle">Open risks by rating</div>${RA.bandBars(RA.BANDS.map((b) => ({ label: b.label, color: b.color, n: open.filter((r) => r.rating === b.label).length })))}
        <div class="sx-ctitle" style="margin-top:14px">By category</div>${bars(countBy(open, (r) => r.category).slice(0, 6))}
        <div class="sx-ctitle" style="margin-top:14px">By status</div>${bars(countBy(all, (r) => r.status, ["Open", "Mitigated", "Closed"]).map((x) => Object.assign(x, { color: /closed/i.test(x.label) ? "#98a2b3" : /mitig/i.test(x.label) ? "#22c55e" : "#f5a524" })))}</div></div>
      <div id="${p}box">${allTbl()}</div>
      ${closed.length ? `<details class="sx-more"><summary>Closed risks (${closed.length})</summary>${tbl(closed)}</details>` : ""}`;
    const right = pan(
      [`${open.length} open risk(s): ${vh.length} Very High, ${hi.length} High, ${open.filter((r) => r.rating === "Medium").length} Medium, ${open.filter((r) => r.rating === "Low" || r.rating === "Very Low").length} Low / Very Low.`, `${mit.length} mitigated, ${closed.length} closed.`, sorted[0] ? `Highest: ${sorted[0].id} — ${trunc(sorted[0].title, 80)} (${sorted[0].rating}).` : ""].filter(Boolean),
      (vh.length + hi.length ? [`${vh.length + hi.length} risk(s) rated High or Very High: ${list(vh.concat(hi), 4)}.`] : []).concat(overdue.length ? [`${overdue.length} review date(s) overdue: ${list(overdue, 4)}.`] : []),
      soon.length ? [`${soon.length} risk review(s) due within 14 days: ${list(soon, 4)}.`] : [],
      vh.length + hi.length ? `Priority is to land mitigations on ${list(vh.concat(hi), 3)}; the heat-map shows where risk is concentrated.` : "No High or Very High risks are open; keep reviewing on the agreed dates.");
    const _rag = rate({ veryHigh: vh.length, high: hi.length });
    return { rag: _rag, html: frame(env, "risks", { title: "Risks", position: `${open.length} open · ${vh.length + hi.length} High / Very High`, rag: _rag, left, right }),
      draw(root) { wireFilter(root, p + "heat", p + "box", (b) => { const l = +b.dataset.l, i = +b.dataset.i; return open.filter((r) => r.l === l && r.i === i); }, tbl, (b, s) => `Risks at likelihood ${b.dataset.l} × impact ${b.dataset.i} (${s.length})`, allTbl); } };
  }

  // ----------------------------------------------------------------------------------------------- ISSUES
  function issues(env) {
    if (!env.raid) return missing(env, "issues", "Issues", NO_RAID);
    const asOf = env.all.asOf, per = env.period, p = env.prefix + "is-";
    const all = env.raid.issues.filter((i) => !i.archived), open = all.filter(RA.isOpenIssue), done = all.filter((i) => !RA.isOpenIssue(i));
    const inP = (d) => d && d >= per.start && d <= per.end;
    const raised = all.filter((i) => inP(i.reported)), resolved = all.filter((i) => inP(i.actual));
    const overdue = open.filter((i) => late(i.target, asOf)), vh = open.filter((i) => i.rating === "Very High"), hi = open.filter((i) => i.rating === "High");
    const age = (i) => (i.reported ? Math.max(0, D.diff(i.reported, asOf)) : null);
    const ages = open.map(age).filter((x) => x != null);
    const tbl = (items) => U.table(["ID", "Issue", "Rating", "Status", "Owner", "Reported", "Age", "Target", "Linked risk"], items.map((i) => [`<b>${esc(i.id)}</b>`, esc(trunc(i.title, 110)), RA.ratingChip(i.rating), statusChip(i.status), esc(trunc(i.owner, 26)), dfmt(i.reported), age(i) != null && RA.isOpenIssue(i) ? age(i) + "d" : "—", `<span style="${late(i.target, asOf) && RA.isOpenIssue(i) ? "color:#b42318;font-weight:700" : ""}">${dfmt(i.target)}</span>`, esc(i.originalRisk || "—")]), { wrap: true });
    const sorted = open.slice().sort((a, b) => (b.score || 0) - (a.score || 0));
    const allTbl = () => `<div class="sx-ctitle">Open issues <span class="sx-csub">· ${open.length}, highest rating first</span></div>${open.length ? tbl(sorted) : '<div class="sx-empty-note">No open issues.</div>'}`;
    const rows = [5, 4, 3, 2, 1].map((v) => ({ key: v, label: ["", "Very Low", "Low", "Medium", "High", "Very High"][v] }));
    const cols = [1, 2, 3, 4, 5].map((v) => ({ key: v, label: ["", "Negligible", "Minor", "Moderate", "Major", "Severe"][v] }));
    const cellItems = (r, c) => open.filter((i) => i.p === r.key && i.s === c.key);
    const buckets = [["0–7 days", 0, 7], ["8–14 days", 8, 14], ["15–30 days", 15, 30], ["Over 30 days", 31, 9999]];
    const left = `<div class="sx-stats">${U.stat(open.length, "Open issues")}${U.stat(vh.length + hi.length, "High / Very High")}${U.stat(raised.length, "Raised this period")}${U.stat(resolved.length, "Resolved this period")}${U.stat(overdue.length, "Past target date")}${U.stat(ages.length ? Math.round(ages.reduce((a, b) => a + b, 0) / ages.length) + "d" : "—", "Average age")}</div>
      <div class="sx-grid2"><div><div class="sx-ctitle">Issue heat-map <span class="sx-csub">· priority × severity, click a cell</span></div><div id="${p}mx">${matrix({ id: p + "mxg", rows, cols, corner: "Priority ↓ / Severity →", cell: (r, c) => ({ items: cellItems(r, c), color: LEVEL_COL[RA.bandOf(r.key * c.key)] }) })}</div></div>
      <div><div class="sx-ctitle">Open issues by rating</div>${RA.bandBars(RA.BANDS.map((b) => ({ label: b.label, color: b.color, n: open.filter((i) => i.rating === b.label).length })))}
        <div class="sx-ctitle" style="margin-top:14px">By status</div>${bars(countBy(all, (i) => i.status, ["Open", "Resolution in progress", "Resolved", "Closed"]).map((x) => Object.assign(x, { color: /resolved|closed/i.test(x.label) ? "#22c55e" : /progress/i.test(x.label) ? "#f5a524" : "#ef5b5b" })))}
        <div class="sx-ctitle" style="margin-top:14px">Age of open issues</div>${bars(buckets.map(([l, a, b]) => ({ label: l, n: ages.filter((x) => x >= a && x <= b).length, color: a > 14 ? "#ef5b5b" : a > 7 ? "#f5a524" : "#22c55e" })))}</div></div>
      <div id="${p}box">${allTbl()}</div>
      ${resolved.length ? `<div><div class="sx-ctitle">Resolved in this period (${resolved.length})</div>${tbl(resolved)}</div>` : ""}
      ${done.length ? `<details class="sx-more"><summary>All resolved / closed issues (${done.length})</summary>${tbl(done)}</details>` : ""}`;
    const right = pan(
      [`${open.length} open issue(s): ${vh.length} Very High, ${hi.length} High.`, `${raised.length} raised and ${resolved.length} resolved in the period.`, sorted[0] ? `Highest: ${sorted[0].id} — ${trunc(sorted[0].title, 80)} (${sorted[0].rating}).` : ""].filter(Boolean),
      (vh.length + hi.length ? [`${vh.length + hi.length} issue(s) rated High or Very High: ${list(vh.concat(hi), 4)}.`] : []).concat(overdue.length ? [`${overdue.length} issue(s) are past their target date: ${list(overdue, 4)}.`] : []),
      open.filter((i) => i.notes).slice(0, 3).map((i) => `${i.id}: ${trunc(i.notes, 140)}`),
      overdue.length || vh.length ? "Open issues need active management; overdue items should be re-forecast or escalated." : "Issues are within target dates and no Very High items are open.");
    const _rag = overdue.length && (vh.length || hi.length) ? "Red" : rate({ veryHigh: vh.length, high: hi.length + overdue.length });
    return { rag: _rag, html: frame(env, "issues", { title: "Issues", position: `${open.length} open · ${overdue.length} past target`, rag: _rag, left, right }),
      draw(root) { wireFilter(root, p + "mxg", p + "box", (b) => cellItems(rows[+b.dataset.r], cols[+b.dataset.c]), tbl, (b, s) => `Issues at ${rows[+b.dataset.r].label} priority × ${cols[+b.dataset.c].label} severity (${s.length})`, allTbl); } };
  }

  // ----------------------------------------------------------------------------------------------- ASSUMPTIONS
  function assumptions(env) {
    if (!env.raid) return missing(env, "assumptions", "Assumptions", NO_RAID);
    const asOf = env.all.asOf, p = env.prefix + "as-";
    const all = env.raid.assumptions.filter((a) => !a.archived);
    const stOf = (a) => (/incorrect/i.test(a.status) ? "Confirmed Incorrect" : /correct|confirmed/i.test(a.status) ? "Confirmed Correct" : "Unconfirmed");
    const confOf = (a) => (/^h/i.test(a.confidence) ? "High" : /^m/i.test(a.confidence) ? "Medium" : /^l/i.test(a.confidence) ? "Low" : "Not rated");
    const unconf = all.filter((a) => stOf(a) === "Unconfirmed"), ok = all.filter((a) => stOf(a) === "Confirmed Correct"), bad = all.filter((a) => stOf(a) === "Confirmed Incorrect");
    const overdue = unconf.filter((a) => late(a.due, asOf)), soon = unconf.filter((a) => a.due && a.due >= asOf && a.due <= D.add(asOf, 30));
    const lowConf = unconf.filter((a) => confOf(a) === "Low");
    const tbl = (items) => U.table(["ID", "Assumption", "Confidence", "Status", "Impact if wrong", "Validation action", "Due", "Raised by"], items.map((a) => [`<b>${esc(a.id)}</b>`, esc(trunc(a.desc, 120)), chip(confOf(a), confOf(a) === "Low" ? "red" : confOf(a) === "Medium" ? "amber" : confOf(a) === "High" ? "green" : "grey"), statusChip(stOf(a)), `<span class="sx-dim">${esc(trunc(a.impact, 110))}</span>`, `<span class="sx-dim">${esc(trunc(a.validation, 110))}</span>`, `<span style="${late(a.due, asOf) && stOf(a) === "Unconfirmed" ? "color:#b42318;font-weight:700" : ""}">${dfmt(a.due)}</span>`, esc(trunc(a.by, 22))]), { wrap: true });
    const sorted = unconf.slice().sort((a, b) => ((a.due || "9999") < (b.due || "9999") ? -1 : 1));
    const allTbl = () => `<div class="sx-ctitle">Assumptions still to validate <span class="sx-csub">· ${unconf.length}, earliest validation date first</span></div>${unconf.length ? tbl(sorted) : '<div class="sx-empty-note">Every assumption has been confirmed.</div>'}`;
    const rows = ["Low", "Medium", "High", "Not rated"].map((k) => ({ key: k, label: k + " confidence" })), cols = ["Unconfirmed", "Confirmed Correct", "Confirmed Incorrect"].map((k) => ({ key: k, label: k }));
    const colorOf = (r, c) => (c.key === "Confirmed Correct" ? "#9ad8b0" : c.key === "Confirmed Incorrect" ? "#ef5b5b" : r.key === "Low" ? "#ff9d5c" : r.key === "Medium" ? "#ffd66b" : "#c9e58f");
    const cellItems = (r, c) => all.filter((a) => confOf(a) === r.key && stOf(a) === c.key);
    const tb = [["Overdue", (a) => late(a.due, asOf), "#ef5b5b"], ["Next 7 days", (a) => a.due && a.due >= asOf && a.due <= D.add(asOf, 7), "#f5a524"], ["8–30 days", (a) => a.due && a.due > D.add(asOf, 7) && a.due <= D.add(asOf, 30), "#ffd66b"], ["31–60 days", (a) => a.due && a.due > D.add(asOf, 30) && a.due <= D.add(asOf, 60), "#9ad8b0"], ["Later", (a) => a.due && a.due > D.add(asOf, 60), "#9ad8b0"], ["No date", (a) => !a.due, "#d0d5dd"]];
    const left = `<div class="sx-stats">${U.stat(all.length, "Assumptions")}${U.stat(unconf.length, "Unconfirmed")}${U.stat(ok.length, "Confirmed correct")}${U.stat(bad.length, "Confirmed incorrect")}${U.stat(overdue.length, "Validation overdue")}${U.stat(soon.length, "Due in 30 days")}</div>
      <div class="sx-grid2"><div><div class="sx-ctitle">Confidence × status <span class="sx-csub">· click a cell</span></div><div>${matrix({ id: p + "mxg", rows, cols, rowW: 120, corner: "Confidence ↓ / Status →", cell: (r, c) => ({ items: cellItems(r, c), color: colorOf(r, c) }) })}</div></div>
      <div><div class="sx-ctitle">When validation is due <span class="sx-csub">· unconfirmed only</span></div>${bars(tb.map(([l, fn, col]) => ({ label: l, n: unconf.filter(fn).length, color: col })))}
        <div class="sx-ctitle" style="margin-top:14px">By status</div>${bars([{ label: "Unconfirmed", n: unconf.length, color: "#f5a524" }, { label: "Correct", n: ok.length, color: "#22c55e" }, { label: "Incorrect", n: bad.length, color: "#ef5b5b" }])}</div></div>
      <div id="${p}box">${allTbl()}</div>
      ${ok.length + bad.length ? `<details class="sx-more"><summary>Confirmed assumptions (${ok.length + bad.length})</summary>${tbl(ok.concat(bad))}</details>` : ""}`;
    const right = pan(
      [`${all.length} assumption(s): ${unconf.length} unconfirmed, ${ok.length} confirmed correct, ${bad.length} confirmed incorrect.`, soon.length ? `${soon.length} to be validated in the next 30 days.` : "", ].filter(Boolean),
      (bad.length ? [`${bad.length} assumption(s) proved incorrect (${list(bad, 4)}) — check the impact and re-plan.`] : []).concat(overdue.length ? [`${overdue.length} validation(s) overdue: ${list(overdue, 4)}.`] : []).concat(lowConf.length ? [`${lowConf.length} unconfirmed assumption(s) have Low confidence: ${list(lowConf, 4)}.`] : []),
      sorted.slice(0, 3).filter((a) => a.validation).map((a) => `${a.id}: ${trunc(a.validation, 130)}${a.due ? " (by " + D.fmt(a.due, false) + ")" : ""}`),
      overdue.length || bad.length ? "Assumptions that are overdue or proven wrong could change the plan; validate them first." : "Assumptions are being validated on time.");
    const _rag = bad.length || overdue.length ? "Amber" : "Green";
    return { rag: _rag, html: frame(env, "assumptions", { title: "Assumptions", position: `${unconf.length} unconfirmed · ${overdue.length} overdue`, rag: _rag, left, right }),
      draw(root) { wireFilter(root, p + "mxg", p + "box", (b) => cellItems(rows[+b.dataset.r], cols[+b.dataset.c]), tbl, (b, s) => `${rows[+b.dataset.r].label} · ${cols[+b.dataset.c].label} (${s.length})`, allTbl); } };
  }

  // ----------------------------------------------------------------------------------------------- DEPENDENCIES
  function deps(env) {
    if (!env.raid) return missing(env, "deps", "Dependencies", NO_RAID);
    const asOf = env.all.asOf, p = env.prefix + "dp-";
    const all = env.raid.deps.filter((d) => !d.archived), open = all.filter(RA.isOpenDep), closed = all.filter((d) => !RA.isOpenDep(d));
    const timing = (d) => (!d.required ? "No date" : d.required < asOf ? "Overdue" : d.required <= D.add(asOf, 30) ? "Next 30 days" : d.required <= D.add(asOf, 60) ? "31–60 days" : "60+ days");
    const TM = ["Overdue", "Next 30 days", "31–60 days", "60+ days", "No date"], TC = { Overdue: "#ef5b5b", "Next 30 days": "#f5a524", "31–60 days": "#ffd66b", "60+ days": "#9ad8b0", "No date": "#d0d5dd" };
    const overdue = open.filter((d) => timing(d) === "Overdue"), due30 = open.filter((d) => timing(d) === "Next 30 days");
    const hiP = open.filter((d) => /^(very high|high)$/i.test(d.priority));
    const tbl = (items) => U.table(["ID", "Dependency", "Type", "Priority", "Needed by", "From", "Owner", "Impact if not met"], items.map((d) => [`<b>${esc(d.id)}</b>`, esc(trunc(d.desc, 110)), esc(d.type || "—"), RA.ratingChip(levelOf(d.priority) === "Unrated" ? d.priority : levelOf(d.priority)), `<span style="${timing(d) === "Overdue" && RA.isOpenDep(d) ? "color:#b42318;font-weight:700" : ""}">${d.required ? D.fmt(d.required, false) : "TBC"}</span>`, esc(trunc(d.from, 26)), esc(trunc(d.owner, 24)), `<span class="sx-dim">${esc(trunc(d.impact, 110))}</span>`]), { wrap: true });
    const sorted = open.slice().sort((a, b) => ((a.required || "9999") < (b.required || "9999") ? -1 : 1));
    const allTbl = () => `<div class="sx-ctitle">Open dependencies <span class="sx-csub">· ${open.length}, earliest needed-by date first</span></div>${open.length ? tbl(sorted) : '<div class="sx-empty-note">No open dependencies.</div>'}`;
    const present = LV.concat(["Unrated"]).filter((l) => open.some((d) => levelOf(d.priority) === l));
    const rows = (present.length ? present : LV).map((k) => ({ key: k, label: k })), cols = TM.map((k) => ({ key: k, label: k }));
    const cellItems = (r, c) => open.filter((d) => levelOf(d.priority) === r.key && timing(d) === c.key);
    const left = `<div class="sx-stats">${U.stat(open.length, "Open dependencies")}${U.stat(hiP.length, "High / Very High")}${U.stat(overdue.length, "Overdue")}${U.stat(due30.length, "Needed in 30 days")}${U.stat(closed.length, "Closed")}</div>
      <div class="sx-grid2"><div><div class="sx-ctitle">Priority × when needed <span class="sx-csub">· click a cell</span></div>${matrix({ id: p + "mxg", rows, cols, corner: "Priority ↓ / Needed →", cell: (r, c) => ({ items: cellItems(r, c), color: c.key === "Overdue" ? "#ef5b5b" : TC[c.key] }) })}</div>
      <div><div class="sx-ctitle">Open by type</div>${bars(countBy(open, (d) => d.type).slice(0, 7))}<div class="sx-ctitle" style="margin-top:14px">Open by supplier / delivering team</div>${bars(countBy(open, (d) => d.from).slice(0, 6))}</div></div>
      <div id="${p}box">${allTbl()}</div>
      ${closed.length ? `<details class="sx-more"><summary>Closed dependencies (${closed.length})</summary>${tbl(closed)}</details>` : ""}`;
    const right = pan(
      [`${open.length} open dependenc${open.length === 1 ? "y" : "ies"}, ${due30.length} needed in the next 30 days.`, `${closed.length} closed.`, sorted[0] && sorted[0].required ? `Next needed: ${sorted[0].id} — ${trunc(sorted[0].desc, 70)} (${D.fmt(sorted[0].required, false)}).` : ""].filter(Boolean),
      (overdue.length ? [`${overdue.length} dependenc${overdue.length === 1 ? "y is" : "ies are"} overdue: ${list(overdue, 4)}.`] : []).concat(hiP.length ? [`${hiP.length} High / Very High priority item(s) open: ${list(hiP, 4)}.`] : []),
      due30.slice(0, 3).map((d) => `${d.id}: chase ${trunc(d.from || "supplier", 40)} for ${D.fmt(d.required, false)} — ${trunc(d.impact, 90)}`),
      overdue.length ? "Overdue dependencies are already delaying the plan; escalate to the delivering owners." : due30.length ? "Dependencies due in the next 30 days need checkpoints with their owners." : "No dependency is due imminently.");
    const _rag = overdue.length ? (overdue.some((d) => /^(very high|high)$/i.test(d.priority)) ? "Red" : "Amber") : "Green";
    return { rag: _rag, html: frame(env, "deps", { title: "Dependencies", position: `${open.length} open · ${overdue.length} overdue`, rag: _rag, left, right }),
      draw(root) { wireFilter(root, p + "mxg", p + "box", (b) => cellItems(rows[+b.dataset.r], cols[+b.dataset.c]), tbl, (b, s) => `${rows[+b.dataset.r].label} priority · needed ${cols[+b.dataset.c].label.toLowerCase()} (${s.length})`, allTbl); } };
  }

  // ----------------------------------------------------------------------------------------------- DECISIONS
  function decisions(env) {
    if (!env.raid) return missing(env, "decisions", "Decisions", NO_RAID);
    const asOf = env.all.asOf, per = env.period, p = env.prefix + "dc-";
    const all = (env.raid.decisions || []).filter((d) => !d.archived);
    if (!all.length) return missing(env, "decisions", "Decisions", "No decisions are logged. Add rows to the <b>Decisions</b> sheet of the RAID log (the template <code>ATS_RAID_Log_Template.xlsx</code> has it) and click <b>Refresh data</b>.");
    const stOf = (d) => (/^(pending|proposed|draft|open|awaiting|tbc)/i.test(d.status) ? "Pending" : /^reject/i.test(d.status) ? "Rejected" : /^supersed/i.test(d.status) ? "Superseded" : /^(approved|agreed|accepted|made)/i.test(d.status) ? "Approved" : d.status || "Pending");
    const ST = ["Pending", "Approved", "Rejected", "Superseded"], SC = { Pending: "#f5a524", Approved: "#22c55e", Rejected: "#ef5b5b", Superseded: "#98a2b3" };
    const pending = all.filter((d) => stOf(d) === "Pending"), inP = all.filter((d) => d.decided && d.decided >= per.start && d.decided <= per.end);
    const stale = pending.filter((d) => late(d.review, asOf));
    const tbl = (items) => U.table(["ID", "Decision", "Category", "Status", "Decided", "By / forum", "Owner", "Rationale", "Impact", "Linked"], items.map((d) => [`<b>${esc(d.id)}</b>`, esc(trunc(d.title, 100)), esc(d.category || "—"), statusChip(stOf(d)), dfmt(d.decided), esc(trunc(d.maker, 26)), esc(trunc(d.owner, 22)), `<span class="sx-dim">${esc(trunc(d.rationale, 120))}</span>`, `<span class="sx-dim">${esc(trunc(d.impact, 100))}</span>`, esc(d.linked || "—")]), { wrap: true });
    const pendSorted = pending.slice().sort((a, b) => ((a.review || "9999") < (b.review || "9999") ? -1 : 1));
    const decided = all.filter((d) => stOf(d) !== "Pending").sort((a, b) => ((b.decided || "") < (a.decided || "") ? -1 : 1));
    const cats = countBy(all, (d) => d.category).map((x) => x.label).slice(0, 8);
    const stPresent = ST.filter((s) => all.some((d) => stOf(d) === s));
    const rows = cats.map((k) => ({ key: k, label: k })), cols = stPresent.map((k) => ({ key: k, label: k }));
    const cellItems = (r, c) => all.filter((d) => (d.category || "—") === r.key && stOf(d) === c.key);
    const allTbl = () => `<div class="sx-ctitle">Decisions awaiting approval <span class="sx-csub">· ${pending.length}</span></div>${pending.length ? tbl(pendSorted) : '<div class="sx-empty-note">No decisions are pending.</div>'}<div class="sx-ctitle" style="margin-top:14px">Decisions taken <span class="sx-csub">· most recent first</span></div>${tbl(decided.slice(0, 12))}`;
    // decisions per month, last 6 months
    const months = []; let mk = D.monthKey(asOf); for (let i = 0; i < 6; i++) { months.unshift(mk); mk = D.prevMonthKey(mk); }
    const left = `<div class="sx-stats">${U.stat(all.length, "Decisions logged")}${U.stat(pending.length, "Pending")}${U.stat(all.filter((d) => stOf(d) === "Approved").length, "Approved")}${U.stat(all.filter((d) => stOf(d) === "Rejected").length, "Rejected")}${U.stat(all.filter((d) => stOf(d) === "Superseded").length, "Superseded")}${U.stat(inP.length, "Taken this period")}</div>
      <div class="sx-grid2"><div><div class="sx-ctitle">Category × status <span class="sx-csub">· click a cell</span></div>${matrix({ id: p + "mxg", rows, cols, rowW: 110, corner: "Category ↓ / Status →", cell: (r, c) => ({ items: cellItems(r, c), color: SC[c.key] || TEAL }) })}</div>
      <div><div class="sx-ctitle">Decisions taken per month</div>${bars(months.map((m) => ({ label: D.fmtMonth(m, false), n: all.filter((d) => d.decided && d.decided.slice(0, 7) === m).length, color: TEAL })))}
        <div class="sx-ctitle" style="margin-top:14px">By status</div>${bars(stPresent.map((s) => ({ label: s, n: all.filter((d) => stOf(d) === s).length, color: SC[s] })))}</div></div>
      <div id="${p}box">${allTbl()}</div>`;
    const right = pan(
      [`${all.length} decision(s) logged: ${all.filter((d) => stOf(d) === "Approved").length} approved, ${pending.length} pending.`, inP.length ? `${inP.length} taken this period: ${list(inP, 4)}.` : "No decisions were taken this period.", decided[0] ? `Latest: ${decided[0].id} — ${trunc(decided[0].title, 80)}${decided[0].decided ? " (" + D.fmt(decided[0].decided, false) + ")" : ""}.` : ""].filter(Boolean),
      (pending.length ? [`${pending.length} decision(s) still awaiting approval: ${list(pending, 4)}.`] : []).concat(stale.length ? [`${stale.length} pending decision(s) are past their review date: ${list(stale, 4)}.`] : []),
      pendSorted.slice(0, 3).map((d) => `${d.id}: ${trunc(d.title, 70)}${d.maker ? " — with " + trunc(d.maker, 30) : ""}${d.notes ? " (" + trunc(d.notes, 60) + ")" : ""}`),
      pending.length ? "Pending decisions should be taken at the next forum so work is not blocked." : "All logged decisions have been taken.");
    const _rag = stale.length ? "Amber" : "Green";
    return { rag: _rag, html: frame(env, "decisions", { title: "Decisions", position: `${pending.length} pending · ${inP.length} taken this period`, rag: _rag, left, right }),
      draw(root) { wireFilter(root, p + "mxg", p + "box", (b) => cellItems(rows[+b.dataset.r], cols[+b.dataset.c]), tbl, (b, s) => `${rows[+b.dataset.r].label} · ${cols[+b.dataset.c].label} (${s.length})`, allTbl); } };
  }

  ATS.raidPages = { risks, issues, assumptions, deps, decisions };
  // page descriptors shared by Monthly Council and Weekly Report
  ATS.raidPageList = [
    { id: "risks", label: "Risks" }, { id: "issues", label: "Issues" }, { id: "assumptions", label: "Assumptions" },
    { id: "deps", label: "Dependencies" }, { id: "decisions", label: "Decisions" },
  ].map((x) => ({ id: x.id, label: x.label, raidKind: x.id, render: (env) => ATS.raidPages[x.id](env), ragOf: (env) => (env.raid ? ATS.raidPages[x.id](env).rag || null : null) }));
})();
