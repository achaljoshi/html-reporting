/* ============================================================
   PowerPoint export — builds a native .pptx (editable charts, tables and text) from the data the
   dashboard shows, on the client PowerPoint theme (assets/js/pptx-template.js: masters, layouts, logos).

   How it works (all offline, no screenshots):
     1. PptxGenJS draws the slides — text, native tables, native charts — at the template's geometry.
     2. The slides are then transplanted into the template package (JSZip) so they sit on the template's
        real slide layouts (title slide / content slide, footer logos, slide numbers, green side bar).
        Titles become real title placeholders, tables use the template's table style.
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date, N = ATS.num, K = ATS.kpi;

  // ---- template palette (theme: accent1 008670, dk2 00362C, accent2 FFD100, accent3 FF7F38, accent5 C8102E, accent6 83B6FF)
  const TEAL = "008670", DARK = "00362C", YEL = "FFD100", ORG = "FF7F38", RED = "C8102E", BLUE = "83B6FF", GREY = "7F7F7F", INK = "262626";
  const RAGC = { Green: "00B050", Amber: "FFC000", Red: "FF0000", Grey: "A6A6A6" };
  const RAGL = { Green: "G", Amber: "A", Red: "R", Grey: "–" };
  const TABLE_STYLE = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"; // Medium Style 2 – Accent 1 (the template's table look)
  const LAYOUT = { title: 1, content: 19 }; // slideLayoutN.xml inside the template ("Title Slide", "Column_B")

  const clip = (s, n) => { s = String(s == null ? "" : s).replace(/\s+/g, " ").trim(); return s.length > n ? s.slice(0, n - 1) + "…" : s; };
  const pct = (v) => (v == null ? null : Math.round(v * 1000) / 10);
  const slug = (s) => String(s || "").replace(/[^A-Za-z0-9]+/g, "_").replace(/^_|_$/g, "");
  const ragOf = (label) => (/^(very high|high)$/i.test(label) ? "Red" : /^medium$/i.test(label) ? "Amber" : label ? "Green" : "Grey");

  // ---------------------------------------------------------------- text sizing (no autofit in the file — we size it ourselves)
  function textHeight(paras, wIn, size) {
    const cw = (size * 0.53) / 72; let h = 0;
    paras.forEach((p) => {
      const w = wIn - (p.bullet ? 0.2 : 0);
      h += Math.max(1, Math.ceil((p.text.length * cw) / w)) * (size * 1.28) / 72 + (p.gap || 0.045);
    });
    return h;
  }
  function pickSize(paras, wIn, hIn, sizes) {
    for (const s of sizes) if (textHeight(paras, wIn, s) <= hIn) return s;
    return sizes[sizes.length - 1];
  }
  const bulletParas = (arr, n) => (arr && arr.length ? arr.map((x) => ({ text: clip(x, n || 260), bullet: true })) : [{ text: "Nothing to report." }]);

  function runs(paras, size, color) {
    return paras.map((p, i) => ({ text: p.text, options: Object.assign({ fontSize: size, breakLine: i < paras.length - 1, paraSpaceAfter: 3, lineSpacingMultiple: 1.05 }, color ? { color } : {}, p.bullet ? { bullet: { indent: 13 } } : {}, p.bold ? { bold: true } : {}) }));
  }
  function heading(slide, text, x, y, w) {
    slide.addText(text, { x, y, w, h: 0.28, fontSize: 14, bold: true, margin: 0, valign: "top", isTextBox: true });
  }
  function ragBox(slide, rag, x, y) {
    slide.addShape("rect", { x, y, w: 0.69, h: 0.25, fill: { color: RAGC[rag] || RAGC.Grey }, line: { type: "none" } });
    slide.addText(RAGL[rag] || "–", { x, y, w: 0.69, h: 0.25, align: "center", valign: "middle", fontSize: 10, color: "FFFFFF", bold: true, margin: 0, isTextBox: true });
  }

  // ---------------------------------------------------------------- slide scaffolding
  const progName = (env) => (env.data && env.data.config && env.data.config.programName) || "HMRC ATS";
  const stamp = (env) => `Source: ${progName(env)} · data as of ${D.fmt(env.all.asOf)}${ATS.isSample() ? " · SAMPLE DATA" : ""}`;

  function newSlide(deck, title, sub, opts) {
    opts = opts || {};
    const s = deck.pptx.addSlide();
    deck.layouts.push(opts.layout || "content");
    const tsz = title.length <= 46 ? undefined : title.length <= 58 ? 26 : title.length <= 70 ? 22 : 19;
    const t = [{ text: title, options: Object.assign({ breakLine: !!sub }, tsz ? { fontSize: tsz } : {}) }];
    if (sub) t.push({ text: sub, options: { fontSize: 20, bold: true, color: GREY } });
    s.addText(t, { x: 0.62, y: 0.16, w: 11.72, h: 0.84, margin: 0, valign: "top", objectName: "TPL_TITLE", isTextBox: true });
    if (opts.foot) s.addText(opts.foot, { x: 0.79, y: 6.6, w: 4.6, h: 0.2, fontSize: 7.5, color: GREY, margin: 0, isTextBox: true });
    if (opts.notes) s.addNotes(opts.notes);
    return s;
  }
  const cellText = (t, o) => ({ text: t == null ? "" : String(t), options: Object.assign({ fontSize: 10, valign: "middle", margin: [0.02, 0.05, 0.02, 0.05] }, o || {}) });
  const ragCell = (rag, o) => cellText("●", Object.assign({ color: RAGC[rag] || RAGC.Grey, align: "center", fontSize: 18 }, o || {}));
  const hdrRow = (arr, o) => arr.map((t, i) => cellText(t, Object.assign({ bold: true, align: o && o.center && o.center.includes(i) ? "center" : "left" }, o && o.opts || {})));

  // ---------------------------------------------------------------- charts (native, restyled to the template palette)
  const axisOpts = (extra) => Object.assign({ catAxisLabelFontSize: 8, valAxisLabelFontSize: 8, catAxisLabelColor: "595959", valAxisLabelColor: "595959", valGridLine: { color: "E5E5E5", size: 0.5 }, catGridLine: { style: "none" }, showLegend: true, legendPos: "b", legendFontSize: 8, legendColor: "404040", showTitle: true, titleFontSize: 10, titleColor: INK, titleBold: true }, extra || {});

  function charts(pptx, res, env) {
    const d = res.detail || {}, C = pptx.charts, out = [];
    const bar = (title, labels, series, colors, extra) => out.push({ title, draw: (s, x, y, w, h) => s.addChart(C.BAR, series.map((a) => ({ name: a[0], labels, values: a[1] })), Object.assign({ x, y, w, h, barDir: "col", chartColors: colors, barGapWidthPct: 60, title }, axisOpts(extra))) });
    const line = (title, labels, series, colors, extra) => out.push({ title, draw: (s, x, y, w, h) => s.addChart(C.LINE, series.map((a) => ({ name: a[0], labels, values: a[1] })), Object.assign({ x, y, w, h, chartColors: colors, lineSize: 2, lineDataSymbolSize: 5, title }, axisOpts(extra))) });
    const wk = (we) => D.fmt(we, false);
    switch (res.key) {
      case "coverage": {
        const L = [["High", d.levels.high], ["Medium", d.levels.med], ["Low", d.levels.low]];
        bar("Coverage % vs target", L.map((x) => x[0]), [["Coverage %", L.map((x) => pct(x[1].pct) || 0)], ["Target %", L.map((x) => pct(x[1].target) || 0)]], [TEAL, DARK], { valAxisMaxVal: 100, valAxisLabelFormatCode: '0"%"' });
        bar("Requirements vs test cases", L.map((x) => x[0]), [["Requirements", L.map((x) => x[1].reqs || 0)], ["Test cases", L.map((x) => x[1].tcs || 0)]], [DARK, TEAL]);
        break;
      }
      case "dde":
        if (d.byPri) {
          bar("Raised vs rejected by priority", d.byPri.map((x) => x.pri), [["Raised", d.byPri.map((x) => x.raised)], ["Rejected", d.byPri.map((x) => x.rejected)]], [TEAL, RED]);
          const wks = Object.keys(d.raisedByWeek).sort();
          bar("Defects raised per week", wks.map(wk), [["Raised", wks.map((w) => d.raisedByWeek[w])]], [TEAL], { showLegend: false });
        }
        break;
      case "environment":
        bar("Availability vs target", ["Availability %", "Target %"], [["Uptime", [pct(d.avail), pct(d.req > 0 ? 1 - d.greenAt / d.req : 1)]], ["Downtime", [pct(1 - d.avail), pct(d.req > 0 ? d.greenAt / d.req : 0)]]], [TEAL, RED], { barGrouping: "stacked", valAxisMaxVal: 100 });
        bar("Downtime hours per week", d.series.slice(-10).map((x) => wk(x.we)), [["Downtime (h)", d.series.slice(-10).map((x) => x.down)]], [TEAL], { showLegend: false });
        break;
      case "aging":
        if (d.counts) {
          bar("Aging buckets (working days)", K.BUCKETS, [["Defects", K.BUCKETS.map((b) => d.counts[b])]], [TEAL], { showLegend: false });
          const l = d.list.slice().sort((a, b) => (a.d.raised < b.d.raised ? -1 : 1)).slice(-10);
          bar("Aging per defect (working days)", l.map((x) => x.d.id.replace(/^.*-/, "#")), [["Aging", l.map((x) => x.days)]], [DARK], { showLegend: false });
        }
        break;
      case "milestones":
        if (d.ev) bar("Milestone status", ["Completed", "On schedule", "Overdue", "Completed late"], [["Milestones", [d.completed, d.onSched, d.overdue, d.late]]], [TEAL], { showLegend: false });
        break;
      case "tsr":
        if (d.items && d.items.length) bar("Working days per TSR (limit " + d.sla + " days)", d.items.map((x) => x.t.ref.slice(-12)), [["Working days", d.items.map((x) => x.wd)]], [TEAL], { showLegend: false, valAxisMaxVal: Math.max(d.sla + 2, ...d.items.map((x) => x.wd + 1)) });
        break;
      case "leakage": bar("Defects found in production (" + d.quarter + ")", K.PRIORITIES, [["Defects", d.counts]], [RED], { showLegend: false, valAxisMaxVal: Math.max(5, ...d.counts) }); break;
      case "csat": if (d.scores && d.scores.length) bar("CSAT score", d.scores.map((s) => wk(s.we)), [["Score", d.scores.map((s) => s.score)]], [TEAL], { showLegend: false, valAxisMaxVal: 5 }); break;
      case "automation":
        bar("Automation coverage %", ["Test data", "Test cases"], [["Automated", [pct(d.td) || 0, pct(d.tc) || 0]], ["Not automated", [d.td == null ? 0 : pct(1 - d.td), d.tc == null ? 0 : pct(1 - d.tc)]]], [TEAL, "BFBFBF"], { barDir: "bar", barGrouping: "stacked", valAxisMaxVal: 100 });
        line("Automation trend", d.series.map((s) => wk(s.we)), [["Test data %", d.series.map((s) => pct(s.td))], ["Test cases %", d.series.map((s) => pct(s.tc))]], [TEAL, BLUE]);
        break;
      case "commercial":
        line("Spend, forecast and budget (£)", d.series.map((s) => wk(s.we)), [["Cumulative spend", d.series.map((s) => s.spend)], ["Forecast", d.series.map((s) => s.forecast)], ["Budget", d.series.map(() => d.budget)]], [TEAL, BLUE, RED], { valAxisLabelFormatCode: "£#,##0" });
        break;
      case "resource": line("FTE planned vs actual", d.series.map((s) => wk(s.we)), [["Planned", d.series.map((s) => s.plan)], ["Actual", d.series.map((s) => s.act)]], [DARK, TEAL]); break;
      case "demand": if (d.rows) bar("Planned vs available FTE", d.rows.map((r) => D.fmtMonth(r.month.slice(0, 7), false)), [["Planned", d.rows.map((r) => r.planned)], ["Available", d.rows.map((r) => r.available)]], [TEAL, DARK]); break;
      case "execution": {
        const s = d.series;
        bar("Execution (cumulative)", s.map((x) => wk(x.we)), [["Passed", s.map((x) => x.passed)], ["Failed", s.map((x) => x.failed)], ["Blocked", s.map((x) => x.blocked)]], [TEAL, RED, YEL], { barGrouping: "stacked" });
        bar("Executed per week", s.map((x) => wk(x.we)), [["Executed", s.map((x, i) => (i ? x.executed - s[i - 1].executed : x.executed))]], [BLUE], { showLegend: false });
        break;
      }
      case "raid":
        if (d.riskBands) bar("Open risks and issues by rating", d.riskBands.map((b) => b.label), [["Risks", d.riskBands.map((b) => b.n)], ["Issues", d.issueBands.map((b) => b.n)]], [ORG, DARK]);
        break;
    }
    return out;
  }

  // ---------------------------------------------------------------- KPI slide (template's two-panel pattern)
  const PANEL = { x: 0.79, y: 1.25, w: 11.54, h: 5.28 };
  function panel(slide) {
    slide.addTable([[{ text: "" }, { text: "" }]], { x: PANEL.x, y: PANEL.y, w: PANEL.w, colW: [5.75, 5.79], rowH: PANEL.h, border: { type: "none" } });
  }

  function kpiSlide(deck, res, env, page) {
    const grp = page && page.group, n = page && page.n;
    const title = res.key === "raid" ? `${progName(env)} (RAID Health)` : `${progName(env)}${grp && n ? ` (${grp} – ${n})` : grp && page.id !== "scorecard" ? ` (${grp})` : ""}`;
    const s = newSlide(deck, title, (page && page.label) || res.name, { foot: stamp(env), notes: `${res.name}: ${res.position}. ${res.exec || ""}` });
    panel(s);
    // ---- left: executive commentary + charts
    const lx = 0.95, lw = 5.4;
    heading(s, "Executive Commentary", lx, 1.4, 4.3);
    ragBox(s, res.rag, 5.62, 1.42);
    const para = [{ text: res.position, bold: true, gap: 0.07 }, { text: clip(res.exec || "—", 560) }];
    const size = pickSize(para, lw, 1.45, [11, 10.5, 10, 9]);
    s.addText(runs(para, size), { x: lx, y: 1.78, w: lw, h: 1.45, margin: 0, valign: "top", isTextBox: true });
    const cs = charts(deck.pptx, res, env).slice(0, 2);
    if (cs.length === 1) cs[0].draw(s, lx - 0.05, 3.3, lw + 0.05, 3.1);
    else if (cs.length === 2) { const w = (lw - 0.12) / 2; cs.forEach((c, i) => c.draw(s, lx - 0.05 + i * (w + 0.17), 3.3, w, 3.1)); }
    // ---- right: highlights / concern / actions
    const rx = 6.7, rw = 5.45, top = 1.4, avail = 5.0;
    const secs = [["Key Highlights", bulletParas(res.highlights, 230)], ["Area of Concern", bulletParas(res.concern, 230)], ["Actions Underway", bulletParas(res.actions, 210)]];
    let fs = 12;
    for (const cand of [13, 12, 11, 10.5, 10, 9.5, 9]) { fs = cand; if (secs.reduce((a, x) => a + 0.34 + textHeight(x[1], rw, cand), 0) <= avail) break; }
    const need = secs.map((x) => 0.34 + textHeight(x[1], rw, fs)), tot = need.reduce((a, b) => a + b, 0), extra = Math.max(0, avail - tot) / 3;
    let y = top;
    secs.forEach((x, i) => {
      heading(s, x[0], rx, y, rw);
      s.addText(runs(x[1], fs), { x: rx, y: y + 0.32, w: rw, h: need[i] - 0.3 + extra, margin: 0, valign: "top", isTextBox: true });
      y += need[i] + extra;
    });
    return s;
  }

  // ---------------------------------------------------------------- full-width table slides
  function tableSlide(deck, env, title, sub, rows, colW, o) {
    o = o || {};
    const s = newSlide(deck, title, sub, { foot: stamp(env) });
    s.addTable(rows, { x: 0.74, y: o.y || 1.3, w: colW.reduce((a, b) => a + b, 0), colW, rowH: o.rowH || 0.34, valign: "middle" });
    return s;
  }

  function scorecardSlide(deck, env, label) {
    const ov = env.all.overall;
    const s = newSlide(deck, `${progName(env)} ${label}`, null, { foot: stamp(env) });
    const rows = [hdrRow(["KPI", "Current Position", "RAG", "Commentary"], { center: [2] })].concat(env.all.scorecard.map((r) => [
      cellText(r.name, { fontSize: 10 }), cellText(clip(r.position, 46), { fontSize: 10 }), ragCell(r.rag), cellText(clip(r.exec, 150), { fontSize: 10 })]));
    s.addTable(rows, { x: 0.74, y: 1.0, w: 11.51, colW: [2.2, 2.25, 0.5, 6.56], rowH: 0.36, valign: "middle" });
    s.addText("Overall", { x: 10.9, y: 0.38, w: 0.7, h: 0.25, fontSize: 10, bold: true, align: "right", margin: 0, valign: "middle", isTextBox: true });
    ragBox(s, ov.rag, 11.66, 0.38);
    return s;
  }

  function lessonsSlide(deck, env) {
    const rows = [hdrRow(["Date", "Category", "Lesson Learned", "Improvement Action", "Owner", "Status"])].concat(env.data.lessons.slice(0, 9).map((l) => [D.fmt(l.date, false), l.category, clip(l.lesson, 170), clip(l.action, 170), l.owner, l.status].map((t) => cellText(t || "", { fontSize: 10.5 }))));
    tableSlide(deck, env, "Risks, Issues, and Lessons Learned", "Lessons Learned (Post Transition)", rows, [0.85, 1.2, 4.0, 3.65, 1.05, 0.75], { y: 1.3, rowH: 0.6 });
  }

  // one slide per RAID sheet (open items first), in the template's table style
  function raidKindSlide(deck, env, kind) {
    const raid = env.raid; if (!raid) return;
    const RA = ATS.raid, asOf = env.all.asOf, base = progName(env);
    const t = (v, o) => cellText(v == null ? "" : v, Object.assign({ fontSize: 10 }, o || {}));
    const dt = (d) => (d ? D.fmt(d, false) : "");
    let title, sub, head, rows, colW;
    if (kind === "risks") {
      const open = raid.risks.filter((r) => !r.archived && RA.isOpenRisk(r)).sort((a, b) => (b.score || 0) - (a.score || 0));
      title = "Key Risks (Open Items)"; sub = `${open.length} open · ${open.filter((r) => /High/.test(r.rating)).length} High / Very High`;
      head = hdrRow(["ID", "Risk", "Impact", "Status", "Owner", "Mitigation", "RAG"], { center: [6] }); colW = [0.6, 3.3, 0.95, 0.95, 1.5, 3.5, 0.69];
      rows = open.slice(0, 8).map((r) => [t(r.id, { bold: true }), t(clip(r.title, 100)), t(r.rating), t(r.status), t(clip(r.owner, 22)), t(clip(r.mitigation, 150), { fontSize: 9 }), ragCell(ragOf(r.rating))]);
    } else if (kind === "issues") {
      const open = raid.issues.filter((i) => !i.archived && RA.isOpenIssue(i)).sort((a, b) => (b.score || 0) - (a.score || 0));
      title = "Current Issues"; sub = `${open.length} open · ${open.filter((i) => i.target && i.target < asOf).length} past target date`;
      head = hdrRow(["ID", "Issue", "Rating", "Status", "Owner", "Reported", "Target"]); colW = [0.6, 4.3, 0.95, 1.6, 1.8, 1.1, 1.1];
      rows = open.slice(0, 9).map((i) => [t(i.id, { bold: true }), t(clip(i.title, 110)), t(i.rating), t(i.status), t(clip(i.owner, 24)), t(dt(i.reported)), t(dt(i.target))]);
    } else if (kind === "assumptions") {
      const open = raid.assumptions.filter((a) => !a.archived && /unconfirmed/i.test(a.status)).sort((a, b) => ((a.due || "9999") < (b.due || "9999") ? -1 : 1));
      title = "Assumptions"; sub = `${open.length} unconfirmed · ${open.filter((a) => a.due && a.due < asOf).length} validation(s) overdue`;
      head = hdrRow(["ID", "Assumption", "Confidence", "Impact if incorrect", "Validation action", "Due"]); colW = [0.6, 3.6, 1.0, 2.7, 2.6, 1.0];
      rows = open.slice(0, 8).map((a) => [t(a.id, { bold: true }), t(clip(a.desc, 120)), t(a.confidence), t(clip(a.impact, 90), { fontSize: 9 }), t(clip(a.validation, 90), { fontSize: 9 }), t(dt(a.due))]);
    } else if (kind === "deps") {
      const open = raid.deps.filter((d) => !d.archived && RA.isOpenDep(d)).sort((a, b) => ((a.required || "9999") < (b.required || "9999") ? -1 : 1));
      title = "Dependencies"; sub = `${open.length} open · ${open.filter((d) => d.required && d.required < asOf).length} overdue`;
      head = hdrRow(["ID", "Dependency", "Priority", "Needed by", "From", "Owner"]); colW = [0.6, 4.7, 1.0, 1.1, 2.1, 2.0];
      rows = open.slice(0, 9).map((d) => [t(d.id, { bold: true }), t(clip(d.desc, 120)), t(d.priority), t(d.required ? dt(d.required) : "TBC"), t(clip(d.from, 28)), t(clip(d.owner, 26))]);
    } else {
      const all = (raid.decisions || []).filter((d) => !d.archived);
      if (!all.length) return;
      const pend = all.filter(RA.isPendingDecision), done = all.filter((d) => !RA.isPendingDecision(d)).sort((a, b) => ((b.decided || "") < (a.decided || "") ? -1 : 1));
      title = "Decisions"; sub = `${pend.length} pending · ${all.length} logged`;
      head = hdrRow(["ID", "Decision", "Status", "Decided", "By / forum", "Owner"]); colW = [0.7, 4.7, 1.1, 1.1, 2.0, 1.9];
      rows = pend.concat(done).slice(0, 9).map((d) => [t(d.id, { bold: true }), t(clip(d.title, 110)), t(d.status), t(dt(d.decided)), t(clip(d.maker, 26)), t(clip(d.owner, 24))]);
    }
    if (!rows.length) rows = [[t("–"), t("Nothing open")].concat(colW.slice(2).map(() => t("")))];
    tableSlide(deck, env, `${base} – ${title}`, sub, [head].concat(rows), colW, { rowH: 0.46, y: 1.3 });
  }

  // KPI-specific detail table (shown on its own slide after the KPI slide)
  function detailSlide(deck, res, env, page) {
    const d = res.detail || {}, base = progName(env);
    if (res.key === "milestones" && d.ev && d.ev.length) {
      const rows = [hdrRow(["Milestone", "Due", "Status", "RAG"], { center: [3] })].concat(d.ev.slice(0, 12).map((e) => [cellText(clip(e.m.name, 80)), cellText(D.fmt(e.m.due, false)), cellText(e.status), ragCell(e.rag)]));
      tableSlide(deck, env, `${base} (Milestone detail)`, "Test Milestone Delivery", rows, [6.2, 1.6, 2.7, 1.0], { rowH: 0.36 });
    } else if (res.key === "readiness" && d.rows && d.rows.length) {
      const areas = K.READINESS_AREAS, w = (11.51 - 3.0 - 0.8) / areas.length;
      const rows = [hdrRow(["Topic", "Overall"].concat(areas.map((a) => clip(a.label, 22))), { center: [1].concat(areas.map((a, i) => i + 2)), opts: { fontSize: 8.5 } })].concat(d.rows.map((r) => [cellText(clip(r.topic, 44), { bold: true }), ragCell(r.overall)].concat(areas.map((a) => ragCell((r.areas[a.key] || {}).rag)))));
      tableSlide(deck, env, `${base} (Readiness)`, "SIT Readiness Dashboard", rows, [3.0, 0.8].concat(areas.map(() => w)), { rowH: 0.45 });
    } else if (res.key === "dde" && d.open && d.open.length) {
      const rows = [hdrRow(["Open defect", "Priority", "Owner", "TCs blocked"], { center: [3] })].concat(d.open.slice(0, 12).map((x) => [cellText(clip(x.id + " — " + x.summary, 90)), cellText(x.priority), cellText(clip(x.owner, 22)), cellText(String(x.blocked), { align: "center" })]));
      tableSlide(deck, env, `${base} (Defect detail)`, "Open Defects", rows, [7.0, 1.3, 2.2, 1.0], { rowH: 0.34 });
    }
  }

  function lookSlide(deck, env) {
    const rows = ATS.lookahead(env, 28).slice(0, 14);
    tableSlide(deck, env, `${progName(env)} Look-ahead`, `${D.fmt(D.add(env.all.asOf, 1), false)} – ${D.fmt(D.add(env.all.asOf, 28))}`,
      [hdrRow(["Date", "Type", "What", "Where"])].concat(rows.map((r) => [D.fmt(r.date, false), r.kind, clip(r.item, 95), clip(r.where, 52)].map((t) => cellText(t)))), [1.2, 1.2, 6.3, 2.81], { rowH: 0.33 });
  }

  function historySlide(deck, env, kind) {
    const isMonth = kind === "month";
    const keys = (isMonth ? K.months(env.data) : K.weeks(env.data).slice(-10)).filter((k) => k <= env.period.key);
    const cols = keys.map((k) => ({ k, res: K.computeAll(env.data, isMonth ? K.monthPeriod(k) : K.weekPeriod(k), { raid: env.raid, withPrev: false }) }));
    const rows = [hdrRow(["KPI"].concat(cols.map((c) => (isMonth ? D.fmtMonth(c.k, false) : D.fmt(c.k, false)))), { center: cols.map((c, i) => i + 1) })]
      .concat([[cellText("Overall (declared)", { bold: true })].concat(cols.map((c) => ragCell(c.res.overall.rag)))])
      .concat(K.KPI_META.filter((m) => !m.extra).map((m) => [cellText(m.name)].concat(cols.map((c) => ragCell(c.res.byKey[m.key].rag)))));
    tableSlide(deck, env, `${progName(env)} ${isMonth ? "RAG History by Month" : "RAG History by Week"}`, null, rows, [3.4].concat(cols.map(() => (11.51 - 3.4) / Math.max(cols.length, 1))), { y: 0.95, rowH: 0.33 });
  }

  // ---------------------------------------------------------------- title + agenda
  function titleSlide(deck, env, kind) {
    const s = deck.pptx.addSlide(); deck.layouts.push("title");
    s.addText(progName(env), { x: 0.68, y: 2.07, w: 12.22, h: 1.23, margin: 0, objectName: "TPL_TITLE", isTextBox: true });
    s.addText((kind === "month" ? "Monthly Unified Quality Council" : "Weekly Delivery & Quality Report") + " - " + (kind === "month" ? D.fmtMonth(env.period.key) : "w/e " + D.fmt(env.period.key, false)), { x: 0.68, y: 3.38, w: 9.33, h: 0.9, margin: 0, objectName: "TPL_SUB", isTextBox: true });
    s.addText(D.fmt(env.period.end).toUpperCase() + (ATS.isSample() ? "  ·  SAMPLE DATA" : ""), { x: 0.68, y: 4.73, w: 8, h: 0.47, fontSize: 20, margin: 0, color: "0B0C0C", isTextBox: true });
  }
  function agendaSlide(deck, items) {
    const s = newSlide(deck, "Agenda", null, {});
    s.addText(items.map((t, i) => ({ text: `${i + 1}.   ${t}`, options: { breakLine: i < items.length - 1, paraSpaceAfter: 10 } })), { x: 0.74, y: 1.6, w: 9, h: 3, fontSize: 14, margin: 0, valign: "top", isTextBox: true });
  }

  // ---------------------------------------------------------------- deck assembly
  function buildMonthlyOrWeekly(kind, env, pages) {
    const pptx = new window.PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.title = progName(env) + " — " + env.period.label;
    const deck = { pptx, layouts: [] };
    titleSlide(deck, env, kind);
    if (kind === "month") agendaSlide(deck, ["Risks, Issues, and Lessons Learned", "ATS Performance Scorecard (KPIs / Metrics)", "Commercials + Resourcing", "Deep Dive items (if any)", "AOB"]);
    pages.forEach((p) => {
      if (p.id === "scorecard" || (p.id === "glance" && kind === "week")) scorecardSlide(deck, env, kind === "month" ? "Overall Dashboard" : "Week at a Glance");
      else if (p.id === "history") historySlide(deck, env, kind);
      else if (p.id === "lessons") { if (env.data.lessons.length) lessonsSlide(deck, env); }
      else if (p.id === "look") lookSlide(deck, env);
      else if (p.kpi && env.all.byKey[p.kpi] && !env.all.byKey[p.kpi].empty) {
        const res = env.all.byKey[p.kpi];
        kpiSlide(deck, res, env, p); detailSlide(deck, res, env, p);
      } else if (p.raidKind) raidKindSlide(deck, env, p.raidKind);
    });
    return deck;
  }

  // ---------------------------------------------------------------- transplant into the template package
  const P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main", A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
  const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  const strip = (e) => { while (e.firstChild) e.removeChild(e.firstChild); while (e.attributes.length) e.removeAttribute(e.attributes[0].name); };

  function postProcessSlide(xml, kind, slideNo) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const all = (tag) => Array.prototype.slice.call(doc.getElementsByTagName(tag));
    // titles -> real placeholders
    all("p:sp").forEach((sp) => {
      const nv = sp.getElementsByTagName("p:cNvPr")[0]; if (!nv) return;
      const nm = nv.getAttribute("name"); if (nm !== "TPL_TITLE" && nm !== "TPL_SUB") return;
      nv.setAttribute("name", nm === "TPL_TITLE" ? "Title" : "Subtitle");
      const cs = sp.getElementsByTagName("p:cNvSpPr")[0]; cs.removeAttribute("txBox");
      while (cs.firstChild) cs.removeChild(cs.firstChild);
      const lock = doc.createElementNS(A_NS, "a:spLocks"); lock.setAttribute("noGrp", "1"); cs.appendChild(lock);
      const nvPr = sp.getElementsByTagName("p:nvPr")[0]; const ph = doc.createElementNS(P_NS, "p:ph");
      ph.setAttribute("type", nm === "TPL_TITLE" ? "ctrTitle" : "subTitle"); if (nm === "TPL_SUB") ph.setAttribute("idx", "1");
      nvPr.appendChild(ph);
      const bp = sp.getElementsByTagName("a:bodyPr")[0]; if (bp) strip(bp);
      // PptxGenJS stamps every run black — drop that so the layout's title colour/size applies (explicit colours stay)
      Array.prototype.slice.call(sp.getElementsByTagName("a:srgbClr")).forEach((c) => { if (c.getAttribute("val") === "000000" && c.parentNode.localName === "solidFill") { const f = c.parentNode; f.parentNode.removeChild(f); } });
      // inherit fill/geometry look from the layout (keep position)
      const sppr = sp.getElementsByTagName("p:spPr")[0];
      ["a:prstGeom", "a:noFill"].forEach((t) => { const e = sppr.getElementsByTagName(t)[0]; if (e && e.parentNode === sppr) sppr.removeChild(e); });
    });
    // tables -> the template's table style; let the style draw borders/fills
    all("a:tbl").forEach((tbl) => {
      const pr = tbl.getElementsByTagName("a:tblPr")[0];
      const hasText = Array.prototype.some.call(tbl.getElementsByTagName("a:t"), (t) => t.textContent.trim());
      pr.setAttribute("bandRow", "1"); if (hasText) pr.setAttribute("firstRow", "1"); else pr.removeAttribute("firstRow");
      Array.prototype.slice.call(pr.getElementsByTagName("a:tableStyleId")).forEach((e) => pr.removeChild(e));
      const id = doc.createElementNS(A_NS, "a:tableStyleId"); id.textContent = TABLE_STYLE; pr.appendChild(id);
      ["a:lnL", "a:lnR", "a:lnT", "a:lnB"].forEach((t) => Array.prototype.slice.call(tbl.getElementsByTagName(t)).forEach((e) => e.parentNode.removeChild(e)));
    });
    // slide number placeholder (content slides)
    if (kind !== "title") {
      const tree = doc.getElementsByTagName("p:spTree")[0];
      const frag = new DOMParser().parseFromString(`<root xmlns:p="${P_NS}" xmlns:a="${A_NS}"><p:sp><p:nvSpPr><p:cNvPr id="9001" name="Slide Number Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="18"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{3D3F7771-BA92-784A-B6DD-D9CB8E3BB7BB}" type="slidenum"><a:rPr lang="en-GB"/><a:t>${slideNo}</a:t></a:fld><a:endParaRPr lang="en-GB"/></a:p></p:txBody></p:sp></root>`, "application/xml");
      tree.appendChild(doc.importNode(frag.documentElement.firstChild, true));
    }
    return new XMLSerializer().serializeToString(doc);
  }

  async function assemble(deck, titleText) {
    if (!window.JSZip) throw new Error("JSZip missing");
    const tplB64 = window.ATS_PPTX_TEMPLATE;
    if (!tplB64) throw new Error("PowerPoint template (assets/js/pptx-template.js) is missing — run tools/build_pptx_template.py");
    const gen = await window.JSZip.loadAsync(await deck.pptx.write({ outputType: "arraybuffer" }));
    const tpl = await window.JSZip.loadAsync(tplB64, { base64: true });
    const text = (z, n) => z.file(n).async("string");

    const slideFiles = Object.keys(gen.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a, b) => parseInt(a.match(/(\d+)\.xml/)[1], 10) - parseInt(b.match(/(\d+)\.xml/)[1], 10));
    const ct = [], presRels = [], sldIds = [];
    for (let i = 0; i < slideFiles.length; i++) {
      const no = i + 1, kind = deck.layouts[i] || "content", base = slideFiles[i].replace(/^ppt\/slides\//, "");
      tpl.file("ppt/slides/" + base, postProcessSlide(await text(gen, slideFiles[i]), kind, no));
      let rels = await text(gen, "ppt/slides/_rels/" + base + ".rels");
      rels = rels.replace(/<Relationship\b[^>]*relationships\/notesSlide"[^>]*\/>/g, "").replace(/<Relationship\b[^>]*relationships\/notesSlide[^>]*\/>/g, "");
      rels = rels.replace(/(<Relationship\b[^>]*relationships\/slideLayout"[^>]*Target=")[^"]*(")/, `$1../slideLayouts/slideLayout${LAYOUT[kind]}.xml$2`);
      rels = rels.replace(/(Target=")\.\.\/slideLayouts\/slideLayout\d+\.xml(")/, `$1../slideLayouts/slideLayout${LAYOUT[kind]}.xml$2`);
      rels = rels.replace(/Target="\.\.\/media\/([^"]+)"/g, 'Target="../media/gen_$1"');
      tpl.file("ppt/slides/_rels/" + base + ".rels", rels);
      ct.push(`<Override PartName="/ppt/slides/${base}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`);
      presRels.push(`<Relationship Id="rIdGen${no}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${base}"/>`);
      sldIds.push(`<p:sldId id="${256 + i}" r:id="rIdGen${no}"/>`);
    }
    // charts, embedded workbooks, media
    for (const n of Object.keys(gen.files)) {
      if (gen.files[n].dir) continue;
      if (/^ppt\/charts\//.test(n) || /^ppt\/embeddings\//.test(n)) tpl.file(n, await gen.file(n).async("uint8array"));
      else if (/^ppt\/media\//.test(n)) tpl.file(n.replace(/^ppt\/media\//, "ppt/media/gen_"), await gen.file(n).async("uint8array"));
    }
    const genCt = await text(gen, "[Content_Types].xml");
    (genCt.match(/<Override\b[^>]*PartName="\/ppt\/charts\/[^"]+"[^>]*\/>/g) || []).forEach((o) => ct.push(o));
    let tct = await text(tpl, "[Content_Types].xml");
    (genCt.match(/<Default\b[^>]*\/>/g) || []).forEach((d) => { const ext = (d.match(/Extension="([^"]+)"/) || [])[1]; if (ext && !new RegExp('<Default[^>]*Extension="' + ext + '"', "i").test(tct)) tct = tct.replace(/(<Types\b[^>]*>)/, "$1" + d); });
    tpl.file("[Content_Types].xml", tct.replace("</Types>", ct.join("") + "</Types>"));
    // presentation.xml + rels
    let pres = await text(tpl, "ppt/presentation.xml");
    pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, "").replace(/(<p:sldSz\b)/, `<p:sldIdLst>${sldIds.join("")}</p:sldIdLst>$1`);
    tpl.file("ppt/presentation.xml", pres);
    tpl.file("ppt/_rels/presentation.xml.rels", (await text(tpl, "ppt/_rels/presentation.xml.rels")).replace("</Relationships>", presRels.join("") + "</Relationships>"));
    // document title
    tpl.file("docProps/core.xml", (await text(tpl, "docProps/core.xml")).replace(/<dc:title>[\s\S]*?<\/dc:title>/, "<dc:title>" + String(titleText).replace(/[<>&]/g, "") + "</dc:title>"));
    return tpl.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation", compression: "DEFLATE" });
  }

  function save(blob, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  ATS.buildDeckBlob = async function (kind, env, pages) {
    const deck = buildMonthlyOrWeekly(kind, env, pages);
    return assemble(deck, progName(env) + " — " + env.period.label);
  };

  ATS.exportPptx = async function (kind, env, pages) {
    if (!window.PptxGenJS) { ATS.toast("PowerPoint library missing", "err"); return; }
    try {
      ATS.toast("Building PowerPoint on the client template…");
      const blob = await ATS.buildDeckBlob(kind, env, pages);
      const name = (kind === "month" ? "Monthly_Council_" : "Weekly_Report_") + (slug(ATS.store.current) ? slug(ATS.store.current) + "_" : "") + env.period.key + ".pptx";
      save(blob, name);
      ATS.toast("Saved " + name + " — template theme, editable charts", "ok");
    } catch (e) { console.error(e); ATS.toast("Export failed: " + e.message, "err"); }
  };

  // ---------------- programme (all portfolios) export
  ATS.exportProgrammePptx = async function (last) {
    if (!window.PptxGenJS) { ATS.toast("PowerPoint library missing", "err"); return; }
    try {
      ATS.toast("Building PowerPoint on the client template…");
      const { rows, period, agg: a } = last;
      const pptx = new window.PptxGenJS(); pptx.layout = "LAYOUT_WIDE";
      const deck = { pptx, layouts: [] };
      const prog = (rows[0] && rows[0].p.kpi.config.programme) || "HMRC ATS";
      const env = { data: { config: { programName: prog } }, all: { asOf: period.end } };
      const t = pptx.addSlide(); deck.layouts.push("title");
      t.addText(prog, { x: 0.68, y: 2.07, w: 12.22, h: 1.23, margin: 0, objectName: "TPL_TITLE", isTextBox: true });
      t.addText("Programme Overview - all portfolios - " + period.label, { x: 0.68, y: 3.38, w: 9.33, h: 0.9, margin: 0, objectName: "TPL_SUB", isTextBox: true });
      t.addText((ATS.isSample() ? "SAMPLE DATA" : D.fmt(period.end).toUpperCase()), { x: 0.68, y: 4.73, w: 8, h: 0.47, fontSize: 20, margin: 0, color: "0B0C0C", isTextBox: true });
      const RANKS = { Grey: 0, Green: 1, Amber: 2, Red: 3 };
      const sevOf = (r) => RANKS[r.all.overall.rag] * 100 + r.all.scorecard.filter((x) => x.rag === "Red").length * 10 + r.all.scorecard.filter((x) => x.rag === "Amber").length;
      const sorted = rows.slice().sort((x, y) => sevOf(y) - sevOf(x) || x.name.localeCompare(y.name));
      const short = (m) => clip(m.name.replace(/ \(.*\)/, ""), 11);
      const foot = `Source: ${prog} · ${period.label}${ATS.isSample() ? " · SAMPLE DATA" : ""}`;
      for (let i = 0; i < sorted.length; i += 12) {
        const chunk = sorted.slice(i, i + 12);
        const s = newSlide(deck, `${prog} Portfolio RAG Heat-map`, sorted.length > 12 ? `Portfolios ${i + 1}–${i + chunk.length} of ${sorted.length}` : period.label, { foot });
        const cw = (11.51 - 2.4 - 0.7) / K.KPI_META.length;
        s.addTable([hdrRow(["Portfolio", "Overall"].concat(K.KPI_META.map(short)), { center: [1].concat(K.KPI_META.map((m, j) => j + 2)) , opts: { fontSize: 8 } })].concat(chunk.map((r) => [cellText(r.name, { bold: true, fontSize: 10 }), ragCell(r.all.overall.rag)].concat(K.KPI_META.map((m) => ragCell(r.all.byKey[m.key].rag, { fontSize: 12 }))))),
        { x: 0.74, y: 1.3, w: 11.51, colW: [2.4, 0.7].concat(K.KPI_META.map(() => cw)), rowH: 0.52, valign: "middle" });
      }
      for (let i = 0; i < sorted.length; i += 12) {
        const chunk = sorted.slice(i, i + 12);
        const s = newSlide(deck, `${prog} Portfolio Summary`, `${a.n} portfolios · ${a.rag.Green} Green · ${a.rag.Amber} Amber · ${a.rag.Red} Red`, { foot });
        const body = [hdrRow(["Portfolio", "Overall", "Executed", "Open defects", "Availability", "Overdue milestones", "TSR over SLA", "High RAID", "Main concern"], { center: [1, 2, 3, 4, 5, 6, 7] })].concat(chunk.map((r) => {
          const g = (k) => r.all.byKey[k];
          const conc = r.all.scorecard.filter((x) => x.rag === "Red").concat(r.all.scorecard.filter((x) => x.rag === "Amber"))[0];
          const num = (v) => cellText(String(v), { align: "center", fontSize: 11 });
          return [cellText(r.name, { bold: true }), ragCell(r.all.overall.rag), num(g("execution").empty ? "–" : N.pct(g("execution").detail.execPct)), num(g("dde").empty || !g("dde").detail.open ? 0 : g("dde").detail.open.length),
            num(g("environment").empty ? "–" : N.pct1(g("environment").detail.avail)), num(g("milestones").empty ? "–" : g("milestones").detail.overdue), num(g("tsr").detail && g("tsr").detail.over ? g("tsr").detail.over : 0),
            num(g("raid").empty ? "–" : g("raid").detail.veryHigh + g("raid").detail.high), cellText(conc ? clip(conc.name + ": " + conc.position, 52) : "All Green", { fontSize: 10 })];
        }));
        s.addTable(body, { x: 0.74, y: 1.3, w: 11.51, colW: [1.9, 0.75, 0.9, 0.95, 1.0, 1.15, 0.95, 0.85, 3.06], rowH: 0.5, valign: "middle" });
      }
      const blob = await assemble(deck, prog + " — Programme overview");
      const name = "Programme_Overview_" + period.key + ".pptx";
      save(blob, name);
      ATS.toast("Saved " + name, "ok");
    } catch (e) { console.error(e); ATS.toast("Export failed: " + e.message, "err"); }
  };
})();
