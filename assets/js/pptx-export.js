/* ============================================================
   PowerPoint export — builds a native .pptx (editable charts + text) from the
   same data the dashboard shows. No screenshots, fully offline (PptxGenJS bundle).
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date, N = ATS.num, R = ATS.rag, K = ATS.kpi;

  const NAVY = "0F1C3F", TEAL = "0EA5A0", INK = "101828", SUB = "667085", LINE = "E4E7EC";
  const RAGC = { Green: "22C55E", Amber: "F5A524", Red: "EF4444", Grey: "98A2B3" };
  const FONT = "Calibri";
  const clip = (s, n) => { s = String(s == null ? "" : s); return s.length > n ? s.slice(0, n - 1) + "…" : s; };
  const bullets = (arr, n) => (arr && arr.length ? arr.map((x) => "• " + clip(x, n || 240)).join("\n") : "Nothing to report.");
  const pct = (v) => (v == null ? null : Math.round(v * 1000) / 10);

  function header(pptx, slide, o) {
    slide.background = { color: "FFFFFF" };
    slide.addText(o.eyebrow || "", { x: 0.5, y: 0.28, w: 9.5, h: 0.28, fontFace: FONT, fontSize: 10.5, color: TEAL, bold: true, charSpacing: 2, margin: 0, isTextBox: true });
    slide.addText(o.title, { x: 0.5, y: 0.55, w: 10.2, h: 0.6, fontFace: "Cambria", fontSize: 26, bold: true, color: NAVY, margin: 0, isTextBox: true });
    if (o.sub) slide.addText(o.sub, { x: 0.5, y: 1.15, w: 10.8, h: 0.32, fontFace: FONT, fontSize: 12, color: SUB, margin: 0, isTextBox: true });
    if (o.rag) {
      slide.addShape("ellipse", { x: 12.2, y: 0.38, w: 0.62, h: 0.62, fill: { color: RAGC[o.rag] || RAGC.Grey }, line: { type: "none" } });
      slide.addText(o.rag === "Green" ? "G" : o.rag === "Amber" ? "A" : o.rag === "Red" ? "R" : "–", { x: 12.2, y: 0.38, w: 0.62, h: 0.62, align: "center", valign: "middle", fontFace: FONT, fontSize: 18, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
      slide.addText(o.rag === "Grey" ? "N/A" : o.rag.toUpperCase(), { x: 11.9, y: 1.02, w: 1.22, h: 0.22, align: "center", fontFace: FONT, fontSize: 8.5, bold: true, color: SUB, margin: 0, isTextBox: true });
    }
    slide.addText(o.foot || "", { x: 0.5, y: 7.12, w: 10, h: 0.25, fontFace: FONT, fontSize: 8.5, color: "98A2B3", margin: 0, isTextBox: true });
  }

  function narrative(slide, res, x, y, w) {
    const boxes = [["Key Highlights", bullets(res.highlights, 220), "F0FAF9", "0A7C78"], ["Area of Concern", bullets(res.concern, 220), "FFF7F6", "B42318"], ["Actions Underway", bullets(res.actions, 200), "F5F6FF", "4A4FB8"], ["Executive Commentary", clip(res.exec || "—", 520), "F8FAFC", NAVY]];
    const hs = [1.62, 1.1, 0.95, 1.6];
    let cy = y;
    boxes.forEach((b, i) => {
      slide.addShape("roundRect", { x, y: cy, w, h: hs[i], rectRadius: 0.08, fill: { color: b[2] }, line: { type: "none" } });
      slide.addText(b[0], { x: x + 0.12, y: cy + 0.05, w: w - 0.24, h: 0.24, fontFace: FONT, fontSize: 10.5, bold: true, color: b[3], margin: 0, isTextBox: true });
      slide.addText(b[1], { x: x + 0.12, y: cy + 0.3, w: w - 0.24, h: hs[i] - 0.34, fontFace: FONT, fontSize: 8.6, color: "344054", valign: "top", margin: 0, fit: "shrink", isTextBox: true });
      cy += hs[i] + 0.08;
    });
  }

  function axisOpts(extra) {
    return Object.assign({ catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, catAxisLabelColor: "667085", valAxisLabelColor: "667085", valGridLine: { color: "EEF1F6", size: 0.5 }, catGridLine: { style: "none" }, showLegend: true, legendPos: "b", legendFontSize: 9, legendColor: "475467" }, extra || {});
  }
  function chartTitle(slide, t, x, y) { slide.addText(t, { x, y, w: 5.6, h: 0.26, fontFace: FONT, fontSize: 10.5, bold: true, color: "344054", margin: 0, isTextBox: true }); }

  // ---------------- per-KPI chart builders: return list of {title, draw(slide,x,y,w,h)}
  function charts(pptx, res, env) {
    const d = res.detail || {}, C = pptx.charts, out = [];
    const bar = (title, labels, series, colors, extra) => out.push({ title, draw: (s, x, y, w, h) => s.addChart(C.BAR, series.map((a) => ({ name: a[0], labels, values: a[1] })), Object.assign({ x, y, w, h, barDir: "col", chartColors: colors, barGapWidthPct: 60 }, axisOpts(extra))) });
    const line = (title, labels, series, colors, extra) => out.push({ title, draw: (s, x, y, w, h) => s.addChart(C.LINE, series.map((a) => ({ name: a[0], labels, values: a[1] })), Object.assign({ x, y, w, h, chartColors: colors, lineSize: 2, lineDataSymbolSize: 6 }, axisOpts(extra))) });
    const wk = (we) => D.fmt(we, false);
    switch (res.key) {
      case "coverage": {
        const L = [["High", d.levels.high], ["Medium", d.levels.med], ["Low", d.levels.low]];
        bar("Coverage % vs target", L.map((x) => x[0]), [["Coverage %", L.map((x) => pct(x[1].pct) || 0)], ["Target %", L.map((x) => pct(x[1].target) || 0)]], [TEAL, NAVY], { valAxisMaxVal: 100, valAxisLabelFormatCode: '0"%"' });
        bar("Requirements vs test cases", L.map((x) => x[0]), [["Requirements", L.map((x) => x[1].reqs || 0)], ["Test cases", L.map((x) => x[1].tcs || 0)]], [NAVY, TEAL]);
        break;
      }
      case "dde":
        if (d.byPri) {
          bar("Raised vs rejected by priority", d.byPri.map((x) => x.pri), [["Raised", d.byPri.map((x) => x.raised)], ["Rejected", d.byPri.map((x) => x.rejected)]], [TEAL, "EF4444"]);
          const wks = Object.keys(d.raisedByWeek).sort();
          bar("Defects raised per week", wks.map(wk), [["Raised", wks.map((w) => d.raisedByWeek[w])]], [TEAL], { showLegend: false });
        }
        break;
      case "environment":
        bar("Availability vs target", ["Availability %", "Target %"], [["Uptime", [pct(d.avail), pct(d.req > 0 ? 1 - d.greenAt / d.req : 1)]], ["Downtime", [pct(1 - d.avail), pct(d.req > 0 ? d.greenAt / d.req : 0)]]], [TEAL, "EF4444"], { barGrouping: "stacked", valAxisMaxVal: 100 });
        bar("Downtime hours per week", d.series.slice(-10).map((x) => wk(x.we)), [["Downtime (h)", d.series.slice(-10).map((x) => x.down)]], [TEAL], { showLegend: false });
        break;
      case "aging":
        if (d.counts) {
          bar("Aging buckets (working days)", K.BUCKETS, [["Defects", K.BUCKETS.map((b) => d.counts[b])]], [TEAL], { showLegend: false });
          const l = d.list.slice().sort((a, b) => (a.d.raised < b.d.raised ? -1 : 1)).slice(-10);
          bar("Aging per defect (working days)", l.map((x) => x.d.id.replace(/^.*-/, "#")), [["Aging", l.map((x) => x.days)]], [NAVY], { showLegend: false });
        }
        break;
      case "milestones":
        if (d.ev) bar("Milestone status", ["Completed", "On schedule", "Overdue", "Completed late"], [["Milestones", [d.completed, d.onSched, d.overdue, d.late]]], [TEAL], { showLegend: false });
        break;
      case "tsr":
        if (d.items && d.items.length) bar("Working days per TSR (limit " + d.sla + " days)", d.items.map((x) => x.t.ref.slice(-12)), [["Working days", d.items.map((x) => x.wd)]], [TEAL], { showLegend: false, valAxisMaxVal: Math.max(d.sla + 2, ...d.items.map((x) => x.wd + 1)) });
        break;
      case "leakage": bar("Defects found in production (" + d.quarter + ")", K.PRIORITIES, [["Defects", d.counts]], ["EF4444"], { showLegend: false, valAxisMaxVal: Math.max(5, ...d.counts) }); break;
      case "csat": if (d.scores && d.scores.length) bar("CSAT score", d.scores.map((s) => wk(s.we)), [["Score", d.scores.map((s) => s.score)]], [TEAL], { showLegend: false, valAxisMaxVal: 5 }); break;
      case "automation":
        bar("Automation coverage %", ["Test data", "Test cases"], [["Automated", [pct(d.td) || 0, pct(d.tc) || 0]], ["Not automated", [d.td == null ? 0 : pct(1 - d.td), d.tc == null ? 0 : pct(1 - d.tc)]]], [TEAL, "D0D5DD"], { barDir: "bar", barGrouping: "stacked", valAxisMaxVal: 100 });
        line("Automation trend", d.series.map((s) => wk(s.we)), [["Test data %", d.series.map((s) => pct(s.td))], ["Test cases %", d.series.map((s) => pct(s.tc))]], [TEAL, "6366F1"]);
        break;
      case "commercial":
        line("Spend, forecast and budget (£)", d.series.map((s) => wk(s.we)), [["Cumulative spend", d.series.map((s) => s.spend)], ["Forecast", d.series.map((s) => s.forecast)], ["Budget", d.series.map(() => d.budget)]], [TEAL, "6366F1", "EF4444"], { valAxisLabelFormatCode: "£#,##0" });
        break;
      case "resource": line("FTE planned vs actual", d.series.map((s) => wk(s.we)), [["Planned", d.series.map((s) => s.plan)], ["Actual", d.series.map((s) => s.act)]], [NAVY, TEAL]); break;
      case "demand": if (d.rows) bar("Planned vs available FTE", d.rows.map((r) => D.fmtMonth(r.month.slice(0, 7), false)), [["Planned", d.rows.map((r) => r.planned)], ["Available", d.rows.map((r) => r.available)]], [TEAL, NAVY]); break;
      case "execution": {
        const s = d.series;
        bar("Execution (cumulative)", s.map((x) => wk(x.we)), [["Passed", s.map((x) => x.passed)], ["Failed", s.map((x) => x.failed)], ["Blocked", s.map((x) => x.blocked)]], [TEAL, "EF4444", "F5A524"], { barGrouping: "stacked" });
        bar("Executed per week", s.map((x) => wk(x.we)), [["Executed", s.map((x, i) => (i ? x.executed - s[i - 1].executed : x.executed))]], ["6366F1"], { showLegend: false });
        break;
      }
      case "raid":
        if (d.riskBands) bar("Open risks by rating", d.riskBands.map((b) => b.label), [["Risks", d.riskBands.map((b) => b.n)]], ["F59E0B"], { showLegend: false });
        break;
    }
    return out;
  }

  function tableForKpi(pptx, slide, res, env, x, y, w) {
    const d = res.detail || {}, hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 9 } });
    const ragCell = (r) => ({ text: r === "Grey" ? "–" : r, options: { fill: { color: RAGC[r] || RAGC.Grey }, color: "FFFFFF", bold: true, align: "center", fontSize: 8.5 } });
    if (res.key === "milestones" && d.ev) {
      const rows = [[hdr("Milestone"), hdr("Due"), hdr("Status"), hdr("RAG")]].concat(d.ev.slice(0, 9).map((e) => [{ text: clip(e.m.name, 42), options: { fontSize: 8.5 } }, { text: D.fmt(e.m.due, false), options: { fontSize: 8.5 } }, { text: e.status, options: { fontSize: 8.5 } }, ragCell(e.rag)]));
      slide.addTable(rows, { x, y, w, colW: [w * 0.5, w * 0.17, w * 0.2, w * 0.13], border: { type: "solid", color: LINE, pt: 0.5 }, fontFace: FONT, rowH: 0.24 });
      return true;
    }
    if (res.key === "readiness" && d.rows) {
      const areas = K.READINESS_AREAS;
      const rows = [[hdr("Topic"), hdr("Overall")].concat(areas.map((a) => hdr(clip(a.label, 14))))].concat(d.rows.map((r) => [{ text: clip(r.topic, 40), options: { fontSize: 9, bold: true } }, ragCell(r.overall)].concat(areas.map((a) => ragCell((r.areas[a.key] || {}).rag)))));
      slide.addTable(rows, { x, y, w, border: { type: "solid", color: LINE, pt: 0.5 }, fontFace: FONT, rowH: 0.32 });
      return true;
    }
    if (res.key === "raid" && d.topRisks) {
      const rows = [[hdr("ID"), hdr("Risk"), hdr("Rating")]].concat(d.topRisks.slice(0, 6).map((r) => [{ text: r.id, options: { fontSize: 8.5, bold: true } }, { text: clip(r.title, 90), options: { fontSize: 8.5 } }, { text: r.rating, options: { fontSize: 8.5 } }]));
      slide.addTable(rows, { x, y, w, colW: [w * 0.1, w * 0.72, w * 0.18], border: { type: "solid", color: LINE, pt: 0.5 }, fontFace: FONT, rowH: 0.28 });
      return true;
    }
    if (res.key === "dde" && d.open && d.open.length) {
      const rows = [[hdr("Open defect"), hdr("Priority"), hdr("Owner"), hdr("TCs blocked")]].concat(d.open.slice(0, 6).map((x2) => [{ text: clip(x2.id + " — " + x2.summary, 70), options: { fontSize: 8.5 } }, { text: x2.priority, options: { fontSize: 8.5 } }, { text: clip(x2.owner, 14), options: { fontSize: 8.5 } }, { text: String(x2.blocked), options: { fontSize: 8.5 } }]));
      slide.addTable(rows, { x, y, w, colW: [w * 0.56, w * 0.14, w * 0.16, w * 0.14], border: { type: "solid", color: LINE, pt: 0.5 }, fontFace: FONT, rowH: 0.27 });
      return true;
    }
    return false;
  }

  function kpiSlide(pptx, res, env, eyebrow, titleOverride) {
    const s = pptx.addSlide();
    header(pptx, s, { eyebrow, title: titleOverride || res.name, sub: `${env.period.label} · ${res.position}`, rag: res.rag, foot: `Source: ${env.data.config.programName} · data as of ${D.fmt(env.all.asOf)}${ATS.isSample() ? " · SAMPLE DATA" : ""}` });
    const cs = charts(pptx, res, env);
    const lx = 0.5, lw = 7.65;
    let y = 1.6;
    if (cs.length === 1) { const ch = res.key === "milestones" ? 1.9 : 2.55; chartTitle(s, cs[0].title, lx, y); cs[0].draw(s, lx, y + 0.27, lw, ch); y += ch + 0.4; }
    else if (cs.length >= 2) { const w = (lw - 0.2) / 2; cs.slice(0, 2).forEach((c, i) => { chartTitle(s, c.title, lx + i * (w + 0.2), y); c.draw(s, lx + i * (w + 0.2), y + 0.27, w, 2.55); }); y += 2.95; }
    const hasTable = tableForKpi(pptx, s, res, env, lx, y, lw);
    if (!cs.length && !hasTable) s.addText(res.position, { x: lx, y: 2.2, w: lw, h: 1, fontFace: FONT, fontSize: 22, bold: true, color: NAVY, margin: 0, isTextBox: true });
    narrative(s, res, 8.45, 1.6, 4.4);
    s.addNotes(`${res.name}: ${res.position}. ${res.exec || ""}`);
    return s;
  }

  function scorecardSlide(pptx, env, title, eyebrow) {
    const s = pptx.addSlide();
    const ov = env.all.overall;
    header(pptx, s, { eyebrow, title, sub: `${env.period.label} · ${env.data.config.programName}`, rag: ov.rag, foot: `Source: ${env.data.config.programName}${ATS.isSample() ? " · SAMPLE DATA" : ""}` });
    const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 10 } });
    const rows = [[hdr("KPI"), hdr("Current position"), hdr("RAG"), hdr("Commentary")]].concat(env.all.scorecard.map((r) => [
      { text: r.name, options: { bold: true, fontSize: 9.5 } }, { text: clip(r.position, 48), options: { fontSize: 9 } },
      { text: r.rag === "Grey" ? "N/A" : r.rag, options: { fill: { color: RAGC[r.rag] || RAGC.Grey }, color: "FFFFFF", bold: true, align: "center", fontSize: 9 } },
      { text: clip(r.exec, 150), options: { fontSize: 8.5, color: "475467" } }]));
    s.addTable(rows, { x: 0.5, y: 1.6, w: 12.33, colW: [2.6, 2.9, 0.8, 6.03], border: { type: "solid", color: LINE, pt: 0.5 }, fontFace: FONT, rowH: 0.36, valign: "middle" });
    return s;
  }

  function lessonsSlide(pptx, env) {
    const s = pptx.addSlide();
    header(pptx, s, { eyebrow: "Risks, Issues and Lessons Learned (2/2)", title: "Lessons Learned", sub: env.period.label, foot: "" });
    const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 10 } });
    const rows = [[hdr("Category"), hdr("Lesson learned"), hdr("Improvement action"), hdr("Owner")]].concat(env.data.lessons.slice(0, 8).map((l) => [l.category, clip(l.lesson, 160), clip(l.action, 160), l.owner].map((t) => ({ text: t || "", options: { fontSize: 9 } }))));
    s.addTable(rows, { x: 0.5, y: 1.6, w: 12.33, colW: [1.8, 4.6, 4.6, 1.33], border: { type: "solid", color: LINE, pt: 0.5 }, fontFace: FONT, rowH: 0.4 });
  }

  function lookSlide(pptx, env) {
    const s = pptx.addSlide();
    header(pptx, s, { eyebrow: "Next 4 weeks", title: "Look-ahead", sub: `${D.fmt(D.add(env.all.asOf, 1), false)} – ${D.fmt(D.add(env.all.asOf, 28))}`, foot: "" });
    const rows = ATS.lookahead(env, 28).slice(0, 14);
    const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 10 } });
    s.addTable([[hdr("Date"), hdr("Type"), hdr("What"), hdr("Where")]].concat(rows.map((r) => [D.fmt(r.date, false), r.kind, clip(r.item, 90), clip(r.where, 50)].map((t) => ({ text: t, options: { fontSize: 9 } })))), { x: 0.5, y: 1.6, w: 12.33, colW: [1.2, 1.1, 6.2, 3.83], border: { type: "solid", color: LINE, pt: 0.5 }, fontFace: FONT, rowH: 0.32 });
  }


  function historySlide(pptx, env, kind) {
    const isMonth = kind === "month";
    const keys = (isMonth ? K.months(env.data) : K.weeks(env.data).slice(-10)).filter((k) => k <= env.period.key);
    const cols = keys.map((k) => ({ k, res: K.computeAll(env.data, isMonth ? K.monthPeriod(k) : K.weekPeriod(k), { raid: env.raid, withPrev: false }) }));
    const s = pptx.addSlide();
    header(pptx, s, { eyebrow: isMonth ? "Month-over-month" : "Week-by-week", title: isMonth ? "RAG History by Month" : "RAG History by Week", sub: "Each cell recalculated as of the end of that period", foot: "" });
    const hdr = (t) => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 9, align: "center" } });
    const cell = (r) => ({ text: r === "Grey" ? "–" : r[0], options: { fill: { color: RAGC[r] || RAGC.Grey }, color: "FFFFFF", bold: true, align: "center", fontSize: 9 } });
    const rows = [[Object.assign(hdr("KPI"), { options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 9 } })].concat(cols.map((c) => hdr(isMonth ? D.fmtMonth(c.k, false) : D.fmt(c.k, false))))]
      .concat([[{ text: "Overall (declared)", options: { bold: true, fontSize: 9 } }].concat(cols.map((c) => cell(c.res.overall.rag)))])
      .concat(K.KPI_META.filter((m) => !m.extra).map((m) => [{ text: m.name, options: { fontSize: 9 } }].concat(cols.map((c) => cell(c.res.byKey[m.key].rag)))));
    s.addTable(rows, { x: 0.5, y: 1.6, w: 12.33, colW: [3.2].concat(cols.map(() => (12.33 - 3.2) / cols.length)), border: { type: "solid", color: LINE, pt: 0.5 }, fontFace: FONT, rowH: 0.3 });
  }

  function titleSlide(pptx, env, kind) {
    const s = pptx.addSlide();
    s.background = { color: NAVY };
    s.addText(env.data.config.programName, { x: 0.8, y: 2.2, w: 11.5, h: 0.9, fontFace: "Cambria", fontSize: 38, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
    s.addText(kind === "month" ? "Monthly Unified Quality Council" : "Weekly Delivery & Quality Report", { x: 0.8, y: 3.15, w: 11.5, h: 0.6, fontFace: FONT, fontSize: 22, color: "9FE7E2", margin: 0, isTextBox: true });
    s.addText(env.period.label + (ATS.isSample() ? "  ·  SAMPLE DATA" : ""), { x: 0.8, y: 3.85, w: 11.5, h: 0.4, fontFace: FONT, fontSize: 14, color: "CADCFC", margin: 0, isTextBox: true });
    s.addText("Generated from the ATS Weekly Data workbook", { x: 0.8, y: 6.7, w: 8, h: 0.3, fontFace: FONT, fontSize: 10, color: "7C8FB8", margin: 0, isTextBox: true });
  }

  // pages: the section's page list (we map the pages that make sense in a deck)
  ATS.buildPptx = function (kind, env, pages) {
    const pptx = new window.PptxGenJS();
    pptx.layout = "LAYOUT_WIDE";
    pptx.title = env.data.config.programName + " — " + env.period.label;
    titleSlide(pptx, env, kind);
    if (kind === "month") {
      const a = pptx.addSlide(); header(pptx, a, { eyebrow: "Agenda", title: "Agenda", foot: "" });
      a.addText(["Risks, Issues and Lessons Learned", "ATS Performance Scorecard (KPIs / Metrics)", "Commercials + Resourcing", "Deep Dive items (if any)", "AOB"].map((t) => ({ text: t, options: { bullet: { type: "number" }, breakLine: true } })), { x: 0.8, y: 1.8, w: 9, h: 3, fontFace: FONT, fontSize: 20, color: INK, paraSpaceAfter: 12, isTextBox: true });
    }
    pages.forEach((p) => {
      if (p.id === "scorecard") scorecardSlide(pptx, env, kind === "month" ? "Overall Dashboard" : "Week at a Glance", kind === "month" ? "ATS Performance Scorecard" : "Weekly Report");
      else if (p.id === "glance" && kind === "week") scorecardSlide(pptx, env, "Week at a Glance", "Weekly Report");
      else if (p.id === "history") historySlide(pptx, env, kind);
      else if (p.id === "lessons") { if (env.data.lessons.length) lessonsSlide(pptx, env); }
      else if (p.id === "look") lookSlide(pptx, env);
      else if (p.kpi && env.all.byKey[p.kpi] && !env.all.byKey[p.kpi].empty) kpiSlide(pptx, env.all.byKey[p.kpi], env, p.group + " — " + p.label + (p.n ? " (" + p.n + ")" : ""), p.label === env.all.byKey[p.kpi].name ? null : null);
    });
    return pptx;
  };


  // ---------------- programme (all portfolios) export
  ATS.exportProgrammePptx = async function (last) {
    if (!window.PptxGenJS) { ATS.toast("PowerPoint library missing", "err"); return; }
    try {
      ATS.toast("Building PowerPoint…");
      const { rows, period, agg: a } = last;
      const pptx = new window.PptxGenJS(); pptx.layout = "LAYOUT_WIDE";
      const prog = (rows[0] && rows[0].p.kpi.config.programme) || "ATS Programme";
      const t = pptx.addSlide(); t.background = { color: NAVY };
      t.addText(prog, { x: 0.8, y: 2.2, w: 11.5, h: 0.9, fontFace: "Cambria", fontSize: 38, bold: true, color: "FFFFFF", margin: 0, isTextBox: true });
      t.addText("Programme Overview — all portfolios", { x: 0.8, y: 3.15, w: 11.5, h: 0.6, fontFace: FONT, fontSize: 22, color: "9FE7E2", margin: 0, isTextBox: true });
      t.addText(period.label + (ATS.isSample() ? "  ·  SAMPLE DATA" : ""), { x: 0.8, y: 3.85, w: 11.5, h: 0.4, fontFace: FONT, fontSize: 14, color: "CADCFC", margin: 0, isTextBox: true });
      const hdr = (x, al) => ({ text: x, options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, fontSize: 8.5, align: al || "center" } });
      const cell = (r) => ({ text: r === "Grey" ? "–" : r[0], options: { fill: { color: RAGC[r] || RAGC.Grey }, color: "FFFFFF", bold: true, align: "center", fontSize: 8.5 } });
      const RANKS = { Grey: 0, Green: 1, Amber: 2, Red: 3 };
      const sevOf = (r) => RANKS[r.all.overall.rag] * 100 + r.all.scorecard.filter((x) => x.rag === "Red").length * 10 + r.all.scorecard.filter((x) => x.rag === "Amber").length;
      const sorted = rows.slice().sort((x, y) => sevOf(y) - sevOf(x) || x.name.localeCompare(y.name));
      // heat-map slide(s): 12 portfolios per slide
      for (let i = 0; i < sorted.length; i += 12) {
        const chunk = sorted.slice(i, i + 12);
        const s = pptx.addSlide();
        header(pptx, s, { eyebrow: "Programme overview", title: "Portfolio RAG Heat-map", sub: period.label + (sorted.length > 12 ? ` · portfolios ${i + 1}–${i + chunk.length} of ${sorted.length}` : ""), foot: ATS.isSample() ? "SAMPLE DATA" : "" });
        const body = [[hdr("Portfolio", "left"), hdr("Overall")].concat(K.KPI_META.map((m) => hdr(clip(m.name.replace(/ \(.*\)/, ""), 11))))]
          .concat(chunk.map((r) => [{ text: r.name, options: { bold: true, fontSize: 9 } }, cell(r.all.overall.rag)].concat(K.KPI_META.map((m) => cell(r.all.byKey[m.key].rag)))));
        s.addTable(body, { x: 0.5, y: 1.6, w: 12.33, colW: [2.4, 0.7].concat(K.KPI_META.map(() => (12.33 - 3.1) / K.KPI_META.length)), border: { type: "solid", color: LINE, pt: 0.5 }, fontFace: FONT, rowH: 0.34 });
      }
      // summary table
      for (let i = 0; i < sorted.length; i += 12) {
        const chunk = sorted.slice(i, i + 12);
        const s = pptx.addSlide();
        header(pptx, s, { eyebrow: "Programme overview", title: "Portfolio Summary", sub: `${a.n} portfolios · ${a.rag.Green} Green · ${a.rag.Amber} Amber · ${a.rag.Red} Red`, foot: "" });
        const body = [[hdr("Portfolio", "left"), hdr("Overall"), hdr("Executed"), hdr("Open defects"), hdr("Availability"), hdr("Overdue milestones"), hdr("TSR over SLA"), hdr("High RAID"), hdr("Main concern", "left")]].concat(chunk.map((r) => {
          const g = (k) => r.all.byKey[k];
          const conc = r.all.scorecard.filter((x) => x.rag === "Red").concat(r.all.scorecard.filter((x) => x.rag === "Amber"))[0];
          const num = (v) => ({ text: String(v), options: { align: "center", fontSize: 9 } });
          return [{ text: r.name, options: { bold: true, fontSize: 9 } }, cell(r.all.overall.rag), num(g("execution").empty ? "–" : N.pct(g("execution").detail.execPct)), num(g("dde").empty || !g("dde").detail.open ? 0 : g("dde").detail.open.length), num(g("environment").empty ? "–" : N.pct1(g("environment").detail.avail)), num(g("milestones").empty ? "–" : g("milestones").detail.overdue), num(g("tsr").detail && g("tsr").detail.over ? g("tsr").detail.over : 0), num(g("raid").empty ? "–" : g("raid").detail.veryHigh + g("raid").detail.high), { text: conc ? clip(conc.name + ": " + conc.position, 46) : "All Green", options: { fontSize: 8.5, color: "475467" } }];
        }));
        s.addTable(body, { x: 0.5, y: 1.6, w: 12.33, colW: [2.2, 0.8, 0.9, 1.0, 1.0, 1.2, 1.0, 0.9, 3.33], border: { type: "solid", color: LINE, pt: 0.5 }, fontFace: FONT, rowH: 0.34 });
      }
      await pptx.writeFile({ fileName: "Programme_Overview_" + period.key + ".pptx" });
      ATS.toast("Saved Programme_Overview_" + period.key + ".pptx", "ok");
    } catch (e) { console.error(e); ATS.toast("Export failed: " + e.message, "err"); }
  };

  ATS.exportPptx = async function (kind, env, pages) {
    if (!window.PptxGenJS) { ATS.toast("PowerPoint library missing", "err"); return; }
    try {
      ATS.toast("Building PowerPoint…");
      const pptx = ATS.buildPptx(kind, env, pages);
      const who = String(ATS.store.current || "").replace(/[^A-Za-z0-9]+/g, "_").replace(/^_|_$/g, "");
      const name = (kind === "month" ? "Monthly_Council_" : "Weekly_Report_") + (who ? who + "_" : "") + env.period.key;
      await pptx.writeFile({ fileName: name + ".pptx" });
      ATS.toast("Saved " + name + ".pptx — native, editable charts", "ok");
    } catch (e) { console.error(e); ATS.toast("Export failed: " + e.message, "err"); }
  };
})();
