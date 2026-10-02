/* ============================================================
   Monthly Council — mirrors the council deck (agenda order)
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const U = ATS.ui, esc = ATS.dom.esc;

  const kpiPage = (key, eyebrow, label, n, title) => ({
    id: key, group: null, label, n, kpi: key,
    render: (env) => ATS.pages.kpiSlide(env, key, eyebrow, { title }),
  });

  ATS.monthly = ATS.makeSection({
    id: "m", kind: "month", eyebrow: "Monthly Council", tagline: "Monthly Unified Quality Council",
    pages(env) {
      const pg = [];
      const add = (group, p) => { p.group = group; pg.push(p); };
      add("Overview", { id: "scorecard", label: "Overall Dashboard", render: (e) => ATS.pages.scorecard(e, { mode: e.state.scMode, title: "Overall Dashboard", eyebrow: "ATS Performance Scorecard (KPIs / Metrics)" }) });
      add("Overview", { id: "history", label: "RAG History", render: (e) => ATS.pages.ragHistory(e, { eyebrow: "Month-over-month", title: "RAG History by Month", extra: false }) });
      add("Risks, Issues & Lessons", Object.assign(kpiPage("raid", "Risks, Issues and Lessons Learned (1/2)", "Risks & Issues", "1/2", "Risks & Issues"), {}));
      add("Risks, Issues & Lessons", { id: "lessons", label: "Lessons Learned", n: "2/2", render: (e) => ATS.pages.lessons(e) });
      const monthly = [["coverage", "Risk Based Test Coverage"], ["dde", "Defect Detection"], ["environment", "Test Execution Downtime"], ["aging", "Defect Aging"], ["milestones", "Test Milestone Delivery"], ["tsr", "TSR Impact Assessment"]];
      monthly.forEach(([k, label], i) => add("Monthly Metrics", kpiPage(k, `Monthly Metrics – ${i + 1}/${monthly.length}`, label, `${i + 1}/${monthly.length}`)));
      const quarterly = [["leakage", "Defect Leakage into Production"], ["csat", "Customer Satisfaction (CSAT)"], ["automation", "Automation Coverage"]];
      quarterly.forEach(([k, label], i) => add("Quarterly Metrics", kpiPage(k, `Quarterly Metrics – ${i + 1}/${quarterly.length}`, label, `${i + 1}/${quarterly.length}`)));
      add("Commercials & Delivery", kpiPage("commercial", "Commercials / Portfolio Delivery", "Commercial Performance"));
      add("Commercials & Delivery", kpiPage("execution", "Portfolio Delivery", "Test Execution Progress"));
      add("Commercials & Delivery", kpiPage("resource", "Resourcing", "Resource Position"));
      add("Commercials & Delivery", kpiPage("demand", "Demand Pipeline Forecast", "Demand Forecast"));
      add("Appendix", kpiPage("readiness", "Appendix — SIT Readiness Dashboard", "SIT Readiness"));
      add("Appendix", { id: "checks", label: "Data Quality", render: (e) => ATS.pages.checks(e) });
      return pg;
    },
  });
})();
