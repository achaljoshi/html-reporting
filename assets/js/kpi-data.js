/* ============================================================
   ATS KPI engine — parses ATS_Weekly_Data workbook and computes every KPI
   for ANY week or month ("as of" that period). Schema: tools/SCHEMA_KPI_WORKBOOK.md
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date, N = ATS.num, R = ATS.rag, X = ATS.xl;
  const K = (ATS.kpi = ATS.kpi || {});

  K.PRIORITIES = ["P1 High", "P2 Medium", "P3 Low"];
  K.KPI_META = [
    { key: "coverage", name: "Risk Based Coverage", cadence: "Monthly", group: "Monthly" },
    { key: "leakage", name: "Production Defect Leakage", cadence: "Quarterly", group: "Quarterly" },
    { key: "dde", name: "Defect Detection Effectiveness", cadence: "Monthly", group: "Monthly" },
    { key: "aging", name: "Defect Aging", cadence: "Monthly", group: "Monthly" },
    { key: "automation", name: "Automation Coverage", cadence: "Quarterly", group: "Quarterly" },
    { key: "milestones", name: "Milestone Delivery", cadence: "Monthly", group: "Monthly" },
    { key: "environment", name: "Environment Availability", cadence: "Monthly", group: "Monthly" },
    { key: "tsr", name: "TSR Impact Assessment", cadence: "Monthly", group: "Monthly" },
    { key: "csat", name: "Customer Satisfaction (CSAT)", cadence: "Quarterly", group: "Quarterly" },
    { key: "resource", name: "Resource Position", cadence: "Monthly", group: "Commercial" },
    { key: "demand", name: "Demand Forecast", cadence: "Monthly", group: "Commercial" },
    { key: "commercial", name: "Commercial Performance", cadence: "Monthly", group: "Commercial" },
    { key: "raid", name: "RAID Health", cadence: "Monthly", group: "RAID" },
    { key: "execution", name: "Test Execution Progress", cadence: "Weekly", group: "Delivery", extra: true },
    { key: "readiness", name: "SIT Readiness", cadence: "Weekly", group: "Delivery", extra: true },
  ];
  K.READINESS_AREAS = [
    { key: "req", label: "Requirements & Arch Intent" }, { key: "blockers", label: "Blockers" }, { key: "resources", label: "Resources" },
    { key: "env", label: "Environments" }, { key: "scripts", label: "Test Scripts & Test Data" }, { key: "kt", label: "Knowledge Transfer" },
  ];
  K.metaByKey = {};
  K.KPI_META.forEach((m) => (K.metaByKey[m.key] = m));
  K.metaByName = {};
  K.KPI_META.forEach((m) => (K.metaByName[N.norm(m.name)] = m));

  const TH_DEFAULTS = {
    hoursPerTestDay: 8, tsrSla: 10, turnSla: 5, turnGreen: 0.8, turnAmber: 0.6, rejGreen: 0.05, rejAmber: 0.1,
    dtMGreen: 24, dtMAmber: 48, dtWGreen: 6, dtWAmber: 12, covHigh: 0.98, covMed: 0.95, covLow: 0, covTol: 0.02,
    msWindow: 5, msAmberShare: 0.2, autoTarget: 0.95, csatPass: 3, resAmber: 0.9, comAmber: 0.05, exAmber: 0.05, exRed: 0.2,
  };
  const TH_KEYS = [
    [/^hourspertestday/, "hoursPerTestDay", "n"], [/^tsrsla/, "tsrSla", "n"], [/^defectturnaroundslaworkingdays/, "turnSla", "n"],
    [/^defectturnaroundgreen/, "turnGreen", "r"], [/^defectturnaroundamber/, "turnAmber", "r"],
    [/^defectrejectiongreen/, "rejGreen", "r"], [/^defectrejectionamber/, "rejAmber", "r"],
    [/^downtimemonthlygreen/, "dtMGreen", "n"], [/^downtimemonthlyamber/, "dtMAmber", "n"],
    [/^downtimeweeklygreen/, "dtWGreen", "n"], [/^downtimeweeklyamber/, "dtWAmber", "n"],
    [/^coveragetargethigh/, "covHigh", "r"], [/^coveragetargetmedium/, "covMed", "r"], [/^coveragetargetlow/, "covLow", "r"],
    [/^coverageambertolerance/, "covTol", "r"], [/^milestoneamberwindow/, "msWindow", "n"], [/^milestoneamberifshare/, "msAmberShare", "r"],
    [/^automationtarget/, "autoTarget", "r"], [/^csatpassmark/, "csatPass", "n"], [/^resourceamber/, "resAmber", "r"],
    [/^commercialamber/, "comAmber", "r"], [/^executionamber/, "exAmber", "r"], [/^executionred/, "exRed", "r"],
  ];

  // ============================================================
  // PARSE
  // ============================================================
  K.detect = (wb) => wb.SheetNames.includes("Weekly_Snapshot") && wb.SheetNames.includes("Defect_Log");

  function normPriority(v) {
    const s = N.str(v).toLowerCase();
    if (!s) return "";
    if (/p1|high|critical|urgent/.test(s)) return "P1 High";
    if (/p2|medium|med/.test(s)) return "P2 Medium";
    if (/p3|low/.test(s)) return "P3 Low";
    return N.str(v);
  }
  function normStatus(v) {
    const s = N.str(v).toLowerCase();
    if (!s) return "";
    if (s.startsWith("reject")) return "Rejected";
    if (s.startsWith("clos")) return "Closed";
    if (s.startsWith("resol")) return "Resolved";
    if (/re.?open/.test(s)) return "Re-Open";
    if (s.includes("progress")) return "In Progress";
    if (s.startsWith("defer")) return "Deferred";
    if (s.startsWith("open")) return "Open";
    return N.str(v);
  }
  const yes = (v) => /^(y|yes|true|1)$/i.test(N.str(v));

  function readConfig(wb) {
    const ws = wb.Sheets["Config"];
    const out = { raw: {}, th: Object.assign({}, TH_DEFAULTS) };
    if (!ws) return out;
    const aoa = X.aoa(ws);
    for (const row of aoa) {
      const key = N.norm(row[0]);
      if (!key || key === "setting") continue;
      const val = row[1];
      out.raw[key] = val;
      for (const [re, name, type] of TH_KEYS) {
        if (re.test(key)) {
          const v = type === "r" ? N.ratio(val, null) : N.num(val);
          if (v != null) out.th[name] = v;
          break;
        }
      }
    }
    const g = (re) => { for (const k in out.raw) if (re.test(k)) return out.raw[k]; return null; };
    out.programName = N.str(g(/^programname/)) || "ATS Program";
    out.project = N.str(g(/^projectphase/));
    out.tsrRef = N.str(g(/^tsrreference/));
    out.ptm = N.str(g(/^programmetestmanager/));
    out.tm = N.str(g(/^testmanager/));
    out.startDate = D.toIso(g(/^programmestartdate/));
    out.endDate = D.toIso(g(/^programmeenddate/));
    out.budget = N.num(g(/^totalbudget/));
    out.goLive = D.toIso(g(/^productiongolivedate/));
    out.mode = N.str(g(/^datamode/)).toUpperCase() || "LIVE";
    out.portfolio = N.str(g(/^portfolioname/)).replace(/^<.*>$/, "");
    out.programme = N.str(g(/^programme$/));
    return out;
  }

  K.parse = function (wb) {
    if (!K.detect(wb)) return null;
    const cfg = readConfig(wb);

    // ---- Weekly snapshots
    const snapsRaw = X.table(wb, "Weekly_Snapshot", ["Week Ending", "Planned Test Days", "Downtime Hours", "Overall RAG"]);
    const snapMap = {};
    snapsRaw.forEach((r) => {
      const we = D.toIso(X.field(r, "weekending"));
      if (!we) return;
      snapMap[we] = {
        weekEnding: we,
        testDays: N.num(X.field(r, "plannedtestdays")),
        downtime: N.num(X.field(r, "downtimehours")),
        highReqs: N.num(X.field(r, "highriskrequirements")), highTcs: N.num(X.field(r, "highrisktestcases")), highCov: N.num(X.field(r, "highriskcoveredreqs")),
        medReqs: N.num(X.field(r, "mediumriskrequirements")), medTcs: N.num(X.field(r, "mediumrisktestcases")), medCov: N.num(X.field(r, "mediumriskcoveredreqs")),
        lowReqs: N.num(X.field(r, "lowriskrequirements")), lowTcs: N.num(X.field(r, "lowrisktestcases")), lowCov: N.num(X.field(r, "lowriskcoveredreqs")),
        tdTotal: N.num(X.field(r, "automatabletestdatasetups")), tdDone: N.num(X.field(r, "automatedtestdatasetups")),
        tcTotal: N.num(X.field(r, "automatabletestcases")), tcDone: N.num(X.field(r, "automatedtestcases")),
        csat: N.num(X.field(r, "csatscore")), csatResp: N.num(X.field(r, "csatresponses")),
        resPlan: N.num(X.field(r, "resourcesplanned")), resAct: N.num(X.field(r, "resourcesactual")),
        spend: N.num(X.field(r, "cumulativespend")), forecast: N.num(X.field(r, "forecastatcompletion")),
        rag: R.norm(X.field(r, "overallrag")), summary: N.str(X.field(r, "overallexecutivesummary")),
      };
    });
    const snapshots = Object.keys(snapMap).sort().map((k) => snapMap[k]);

    // ---- Topic progress
    const topicProgress = X.table(wb, "Topic_Progress", ["Week Ending", "Topic", "Phase", "TCs Planned"])
      .map((r) => ({
        weekEnding: D.toIso(X.field(r, "weekending")), topic: N.str(X.field(r, "topic")), phase: N.str(X.field(r, "phase")),
        planned: N.num(X.field(r, "tcsplanned")) || 0, designed: N.num(X.field(r, "tcsdesigned")) || 0,
        passed: N.num(X.field(r, "passed")) || 0, failed: N.num(X.field(r, "failed")) || 0, blocked: N.num(X.field(r, "blocked")) || 0,
        comments: N.str(X.field(r, "comments")),
      }))
      .filter((r) => r.weekEnding && r.topic)
      .sort((a, b) => (a.weekEnding < b.weekEnding ? -1 : a.weekEnding > b.weekEnding ? 1 : 0));

    // ---- Defects
    const defects = X.table(wb, "Defect_Log", ["Defect ID", "Summary", "Priority", "Raised Date", "Status"])
      .map((r) => {
        const sev = N.str(X.field(r, "severity"));
        return {
          id: N.str(X.field(r, "defectid")), summary: N.str(X.field(r, "summary")), topic: N.str(X.field(r, "topic")),
          testType: N.str(X.field(r, "testtype")) || "SIT", priority: normPriority(X.field(r, "priority")), severity: sev,
          sevNum: parseInt(sev, 10) || null, raised: D.toIso(X.field(r, "raiseddate")), resolved: D.toIso(X.field(r, "resolveddate")),
          status: normStatus(X.field(r, "status")), foundIn: /prod/i.test(N.str(X.field(r, "foundin")) + " " + N.str(X.field(r, "testtype"))) ? "Production" : "Testing",
          owner: N.str(X.field(r, "ownerdg", "owner")), blocked: N.num(X.field(r, "tcsblocked")) || 0, triage: N.str(X.field(r, "triagecomments")),
        };
      })
      .filter((d) => d.id && d.raised);

    // ---- TSRs
    const tsrs = X.table(wb, "TSR_Log", ["TSR Ref", "Description", "Received Date", "Returned Date"])
      .map((r) => ({
        ref: N.str(X.field(r, "tsrref")), desc: N.str(X.field(r, "description")), topic: N.str(X.field(r, "topic")),
        received: D.toIso(X.field(r, "receiveddate")), returned: D.toIso(X.field(r, "returneddate")),
        hold: N.num(X.field(r, "holddays")) || 0, status: N.str(X.field(r, "status")) || "Open", value: N.num(X.field(r, "value")),
      }))
      .filter((t) => t.ref && t.received);

    // ---- Milestones
    const milestones = X.table(wb, "Milestones", ["Milestone", "Due Date", "Completed Date", "Critical?"])
      .map((r) => ({
        name: N.str(X.field(r, "milestone")), topic: N.str(X.field(r, "topic")), due: D.toIso(X.field(r, "duedate")),
        forecast: D.toIso(X.field(r, "forecastdate")), completed: D.toIso(X.field(r, "completeddate")),
        owner: N.str(X.field(r, "owner")), critical: yes(X.field(r, "critical")),
      }))
      .filter((m) => m.name && m.due);

    // ---- Readiness
    const AREAS = [
      ["req", "Requirements & Arch Intent", "requirementsarchintent"], ["blockers", "Blockers", "blockers"], ["resources", "Resources", "resources"],
      ["env", "Environments", "environments"], ["scripts", "Test Scripts & Test Data", "testscriptstestdata"], ["kt", "Knowledge Transfer", "knowledgetransfer"],
    ];
    const readiness = X.table(wb, "Readiness", ["Week Ending", "Topic", "Planned SIT Start", "Overall RAG"])
      .map((r) => {
        const areas = {};
        AREAS.forEach(([k, , p]) => {
          areas[k] = { rag: R.norm(r[N.norm(p + "rag")]), note: N.str(r[N.norm(p + "note")]) };
        });
        return {
          weekEnding: D.toIso(X.field(r, "weekending")), topic: N.str(X.field(r, "topic")), sitStart: D.toIso(X.field(r, "plannedsitstart")),
          overall: R.norm(X.field(r, "overallrag")), areas, roadToGreen: N.str(X.field(r, "roadtogreen")),
        };
      })
      .filter((r) => r.weekEnding && r.topic);

    // ---- Commentary
    const commentary = X.table(wb, "Commentary", ["Week Ending", "KPI", "Key Highlights", "Executive Commentary"])
      .map((r) => ({
        weekEnding: D.toIso(X.field(r, "weekending")), kpi: N.str(X.field(r, "kpi")), highlights: N.str(X.field(r, "keyhighlights")),
        concern: N.str(X.field(r, "areaofconcern")), actions: N.str(X.field(r, "actionsunderway")), exec: N.str(X.field(r, "executivecommentary")),
        ragOverride: R.norm(X.field(r, "ragoverride")), positionOverride: N.str(X.field(r, "currentpositionoverride")),
      }))
      .filter((c) => c.weekEnding && c.kpi);

    const resourcing = X.table(wb, "Resourcing", ["Role", "Name", "Organisation", "Allocation"])
      .map((r) => ({
        role: N.str(X.field(r, "role")), name: N.str(X.field(r, "name")), org: N.str(X.field(r, "organisation")),
        alloc: N.ratio(X.field(r, "allocation"), 1), start: D.toIso(X.field(r, "startdate")), end: D.toIso(X.field(r, "enddate")),
        status: N.str(X.field(r, "status")) || "Active",
      }))
      .filter((r) => r.role);
    const demand = X.table(wb, "Demand_Forecast", ["Month", "Planned FTE", "Available FTE"])
      .map((r) => ({ month: D.toIso(X.field(r, "month")), planned: N.num(X.field(r, "plannedfte")), available: N.num(X.field(r, "availablefte")), notes: N.str(r.notes) }))
      .filter((r) => r.month)
      .sort((a, b) => (a.month < b.month ? -1 : 1));
    const lessons = X.table(wb, "Lessons_Learned", ["Date", "Category", "Lesson Learned", "Improvement Action"])
      .map((r) => ({
        date: D.toIso(X.field(r, "date")), category: N.str(X.field(r, "category")), lesson: N.str(X.field(r, "lessonlearned")),
        action: N.str(X.field(r, "improvementaction")), owner: N.str(X.field(r, "owner")), status: N.str(X.field(r, "status")),
      }))
      .filter((r) => r.lesson);
    const holidays = [];
    const hws = wb.Sheets["Bank_Holidays"];
    if (hws) X.aoa(hws).forEach((row) => { const iso = D.toIso(row[0]); if (iso && /^20\d\d-/.test(iso)) holidays.push(iso); });

    return {
      kind: "kpi", config: { portfolio: cfg.portfolio, programme: cfg.programme, programName: cfg.programName, project: cfg.project, tsrRef: cfg.tsrRef, ptm: cfg.ptm, tm: cfg.tm, startDate: cfg.startDate, endDate: cfg.endDate, budget: cfg.budget, goLive: cfg.goLive, mode: cfg.mode },
      th: cfg.th, snapshots, topicProgress, defects, tsrs, milestones, readiness, commentary, resourcing, demand, lessons, holidays,
    };
  };

  K.describe = (d) => `${d.snapshots.length} weeks · ${d.defects.length} defects · ${d.tsrs.length} TSRs · ${d.milestones.length} milestones`;

  // ============================================================
  // Periods
  // ============================================================
  K.weeks = (data) => data.snapshots.map((s) => s.weekEnding);
  K.months = function (data) {
    const set = new Set();
    data.snapshots.forEach((s) => set.add(s.weekEnding.slice(0, 7)));
    return Array.from(set).sort();
  };
  K.weekPeriod = (we) => ({ type: "week", key: we, start: D.add(we, -6), end: we, label: "Week ending " + D.fmt(we), short: "W/E " + D.fmt(we, false) });
  K.monthPeriod = (key) => ({ type: "month", key, start: D.monthStart(key), end: D.monthEnd(key), label: D.fmtMonth(key), short: D.fmtMonth(key, false) });
  K.prevPeriod = (p) => (p.type === "week" ? K.weekPeriod(D.add(p.key, -7)) : K.monthPeriod(D.prevMonthKey(p.key)));

  function holSet(data) { return data._hols || (data._hols = new Set(data.holidays)); }
  function snapAt(data, end) { let s = null; for (const x of data.snapshots) { if (x.weekEnding <= end) s = x; else break; } return s; }
  function snapsIn(data, a, b) { return data.snapshots.filter((x) => x.weekEnding >= a && x.weekEnding <= b); }
  K.snapAt = snapAt;

  const bucket = (days) => (days <= 2 ? "0-2 Days" : days <= 5 ? "3-5 Days" : days <= 10 ? "6-10 Days" : ">10 Days");
  K.BUCKETS = ["0-2 Days", "3-5 Days", "6-10 Days", ">10 Days"];

  // ============================================================
  // KPI computations
  // ============================================================
  function empty(key, msg, extra) {
    const m = K.metaByKey[key];
    return Object.assign({ key, name: m.name, cadence: m.cadence, group: m.group, rag: "Grey", autoRag: "Grey", position: msg || "No data", empty: true, metric: null, highlights: [], concern: [], actions: [], exec: "", detail: {} }, extra || {});
  }
  function base(key, p) {
    const m = K.metaByKey[key];
    return { key, name: m.name, cadence: m.cadence, group: m.group, empty: false, highlights: [], concern: [], actions: [], exec: "", detail: {}, metric: null, rag: "Grey" };
  }
  const pctS = (v) => (v == null ? "—" : Math.round(v * 100) + "%");
  const pct1S = (v) => (v == null ? "—" : (v * 100).toFixed(1) + "%");
  const plural = N.plural;

  function covLevel(s, lvl, th) {
    const reqs = s[lvl + "Reqs"], tcs = s[lvl + "Tcs"], cov = s[lvl + "Cov"];
    const target = { high: th.covHigh, med: th.covMed, low: th.covLow }[lvl];
    const pct = reqs > 0 && cov != null ? cov / reqs : null;
    let rag = "Grey";
    if (pct != null) rag = target <= 0 ? "Green" : pct >= target ? "Green" : pct >= target - th.covTol ? "Amber" : "Red";
    return { reqs, tcs, cov, pct, target, variance: pct != null ? pct - target : null, tpr: reqs > 0 && tcs != null ? tcs / reqs : null, rag };
  }

  function latestTopicRows(data, end) {
    const m = {};
    data.topicProgress.forEach((r) => { if (r.weekEnding <= end) m[r.topic] = r; });
    return Object.values(m);
  }

  function kCoverage(data, p, ctx) {
    const s = snapAt(data, p.end);
    if (!s || s.highReqs == null) return empty("coverage", "No coverage data");
    const th = data.th, lv = { high: covLevel(s, "high", th), med: covLevel(s, "med", th), low: covLevel(s, "low", th) };
    const r = base("coverage", p);
    const rags = [lv.high.rag, lv.med.rag].filter((x) => x !== "Grey");
    r.autoRag = r.rag = rags.length ? R.worst(rags) : "Grey";
    r.position = `${pctS(lv.high.pct)} High / ${pctS(lv.med.pct)} Medium`;
    r.metric = { value: lv.high.pct, text: pctS(lv.high.pct), label: "High-risk coverage", better: "up", fmt: "pct" };
    const totReqs = (lv.high.reqs || 0) + (lv.med.reqs || 0) + (lv.low.reqs || 0), totTcs = (lv.high.tcs || 0) + (lv.med.tcs || 0) + (lv.low.tcs || 0);
    const scope = latestTopicRows(data, p.end).filter((t) => t.planned > 0).sort((a, b) => b.planned - a.planned);
    const scopeTotal = scope.reduce((a, t) => a + t.planned, 0);
    r.detail = { levels: lv, snapWeek: s.weekEnding, totReqs, totTcs, scope, scopeTotal,
      series: data.snapshots.filter((x) => x.weekEnding <= p.end && x.highReqs != null).map((x) => ({ we: x.weekEnding, high: covLevel(x, "high", th).pct, med: covLevel(x, "med", th).pct, low: covLevel(x, "low", th).pct })) };
    r.highlights.push(`The test portfolio comprises ${N.int(totTcs)} test cases covering ${N.int(totReqs)} requirements.`);
    r.highlights.push(`High risk coverage is ${pct1S(lv.high.pct)} (target ${pctS(lv.high.target)}), Medium ${pct1S(lv.med.pct)} (target ${pctS(lv.med.target)}), Low ${pct1S(lv.low.pct)}.`);
    if (scope.length > 1 && scopeTotal > 0) {
      const t = scope[0];
      r.highlights.push(`${t.topic} is the largest area of scope with ${N.int(t.planned)} test cases (${pct1S(t.planned / scopeTotal)}).`);
    }
    if (lv.high.rag === "Amber" || lv.high.rag === "Red") r.concern.push(`High risk coverage is ${Math.abs(lv.high.variance * 100).toFixed(1)} points below the ${pctS(lv.high.target)} target.`);
    if (lv.med.rag === "Amber" || lv.med.rag === "Red") r.concern.push(`Medium risk coverage is ${Math.abs(lv.med.variance * 100).toFixed(1)} points below the ${pctS(lv.med.target)} target.`);
    if (lv.low.pct != null && lv.low.pct < 0.8) r.concern.push(`Low risk coverage is ${pctS(lv.low.pct)} and will continue to build through the test phase.`);
    r.exec = r.rag === "Green"
      ? `Risk-based coverage is on track: ${pctS(lv.high.pct)} of High and ${pctS(lv.med.pct)} of Medium risk requirements are covered, so testing effort is concentrated on the areas of greatest business risk.`
      : `Coverage of ${r.rag === "Red" ? "business-critical" : "higher-risk"} requirements is below target (High ${pctS(lv.high.pct)}, Medium ${pctS(lv.med.pct)}); closing this gap is the priority to protect assurance of the highest-risk functionality.`;
    return r;
  }

  const ofP = (list, pri) => list.filter((d) => d.priority === pri);

  function kDde(data, p, ctx) {
    const all = data.defects.filter((d) => d.raised <= p.end);
    const inPer = all.filter((d) => d.raised >= p.start);
    if (!all.length) return empty("dde", "No defects raised", { detail: { empty: true } });
    const th = data.th, r = base("dde", p);
    const valid = all.filter((d) => d.status !== "Rejected");
    const ft = valid.filter((d) => d.foundIn === "Testing").length, fp = valid.filter((d) => d.foundIn === "Production").length;
    const dde = ft + fp > 0 ? ft / (ft + fp) : null;
    const byPri = K.PRIORITIES.map((pri) => {
      const raised = ofP(inPer, pri), rej = raised.filter((d) => d.status === "Rejected");
      return { pri, raised: raised.length, rejected: rej.length, rejPct: raised.length ? rej.length / raised.length : 0, cumul: ofP(valid, pri).length };
    });
    const totRaised = inPer.length, totRej = inPer.filter((d) => d.status === "Rejected").length;
    const rejPct = totRaised ? totRej / totRaised : 0;
    const rejRag = (x) => (x <= th.rejGreen ? "Green" : x <= th.rejAmber ? "Amber" : "Red");
    r.autoRag = r.rag = rejRag(rejPct);
    const phases = Array.from(new Set(valid.map((d) => d.testType))).join("/") || "testing";
    r.position = `${ft} defect${ft === 1 ? "" : "s"} found in ${phases}` + (fp ? ` · ${fp} escaped to production` : "");
    r.metric = { value: ft, text: String(ft), label: "Defects found in testing", better: null, fmt: "int" };
    const openD = valid.filter((d) => !(d.resolved && d.resolved <= p.end));
    const sev = {}; valid.forEach((d) => { const k = d.severity || "Unclassified"; sev[k] = (sev[k] || 0) + 1; });
    const raisedByWeek = {};
    all.forEach((d) => { const we = D.weekEnding(d.raised); raisedByWeek[we] = (raisedByWeek[we] || 0) + 1; });
    r.detail = { byPri, totRaised, totRej, rejPct, ft, fp, dde, open: openD, sev, rejRag: rejRag(rejPct),
      raisedByWeek, cumulTotal: valid.length, perPeriod: inPer.length };
    const breakdown = K.PRIORITIES.map((pri, i) => ({ pri, n: byPri[i].cumul })).filter((x) => x.n).map((x) => `${x.n} ${x.pri.split(" ")[1]}`).join(", ");
    r.highlights.push(`${plural(ft, "defect")} identified during ${phases} to date (${breakdown || "none"}); ${plural(inPer.length, "defect")} raised in ${p.type === "week" ? "the week" : "the month"}.`);
    r.highlights.push(`${totRej} rejected after investigation (${pct1S(rejPct)} rejection rate)${fp ? "" : "; no defects have escaped to production"}.`);
    const openP1 = openD.filter((d) => d.priority === "P1 High");
    if (openD.length) r.concern.push(`${plural(openD.length, "defect")} remain open${openP1.length ? ` (${openP1.length} High priority)` : ""}, blocking ${N.int(openD.reduce((a, d) => a + d.blocked, 0))} test cases: ${openD.slice(0, 4).map((d) => d.id).join(", ")}.`);
    const owners = Array.from(new Set(openD.map((d) => d.owner).filter(Boolean)));
    if (owners.length) r.actions.push(`Awaiting resolution / fix confirmation from ${owners.join(", ")}.`);
    r.exec = `Defect identification remains effective: ${plural(ft, "defect")} found in testing and ${fp ? plural(fp, "defect") + " in production" : "none in production"}, giving a detection effectiveness of ${dde == null ? "n/a" : pctS(dde)}. The rejection rate is ${pct1S(rejPct)}, which is ${r.rag === "Green" ? "within" : "above"} tolerance.`;
    return r;
  }

  function agingList(data, p, asOf) {
    const hols = holSet(data);
    return data.defects.filter((d) => d.raised <= asOf && d.status !== "Rejected").map((d) => {
      const done = d.resolved && d.resolved <= asOf;
      const end = done ? d.resolved : asOf;
      let days = D.networkdays(d.raised, end, hols);
      if (days < 1) days = 0;
      return { d, days, done: !!done, bucket: bucket(days), met: days <= data.th.turnSla };
    });
  }
  K.agingList = agingList;

  function kAging(data, p, ctx) {
    const asOf = ctx.asOf, list = agingList(data, p, asOf), th = data.th;
    if (!list.length) return empty("aging", "No defects to age");
    const r = base("aging", p);
    const met = list.filter((x) => x.met).length, compliance = met / list.length;
    const resolved = list.filter((x) => x.done), open = list.filter((x) => !x.done);
    const avg = resolved.length ? resolved.reduce((a, x) => a + x.days, 0) / resolved.length : null;
    const counts = {}; K.BUCKETS.forEach((b) => (counts[b] = list.filter((x) => x.bucket === b).length));
    const openOver10 = open.filter((x) => x.days > 10);
    r.autoRag = r.rag = compliance >= th.turnGreen ? "Green" : compliance >= th.turnAmber ? "Amber" : "Red";
    if (openOver10.length && r.rag === "Green") r.autoRag = r.rag = "Amber";
    r.position = `${Math.round(compliance * 100)}% within ${th.turnSla}-day SLA`;
    r.metric = { value: compliance, text: pctS(compliance), label: "Within SLA", better: "up", fmt: "pct" };
    const top = Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0];
    r.detail = { list: list.slice().sort((a, b) => b.days - a.days), counts, met, compliance, avg, open, resolved, openOver10, mostCommon: top, sla: th.turnSla };
    r.highlights.push(`${plural(list.length, "defect")} in scope: ${resolved.length} resolved, ${open.length} open. ${met} (${pctS(compliance)}) are within the ${th.turnSla} working-day SLA.`);
    r.highlights.push(`Aging profile — ${K.BUCKETS.map((b) => `${counts[b]} in ${b}`).join(", ")}${avg != null ? `; average turnaround of resolved defects ${avg.toFixed(1)} working days` : ""}.`);
    if (openOver10.length) r.concern.push(`${plural(openOver10.length, "open defect")} aged beyond 10 working days: ${openOver10.map((x) => `${x.d.id} (${x.days}d)`).join(", ")}.`);
    r.exec = r.rag === "Green"
      ? `The defect backlog is well managed: ${pctS(compliance)} of defects sit within the ${th.turnSla}-day SLA${top ? ` and the largest group is ${top}` : ""}.`
      : `Defect turnaround needs attention: only ${pctS(compliance)} of defects are within the ${th.turnSla}-day SLA${openOver10.length ? `, with ${openOver10.length} open beyond 10 days` : ""}.`;
    return r;
  }

  function kLeakage(data, p, ctx) {
    const q = D.quartersOf ? D.quartersOf(p.end) : D.quarterOf(p.end);
    const qEnd = D.minIso(q.end, p.end);
    const inProd = data.config.goLive && data.config.goLive <= p.end;
    const prod = data.defects.filter((d) => d.foundIn === "Production" && d.raised >= q.start && d.raised <= qEnd && d.status !== "Rejected");
    const c = K.PRIORITIES.map((pri) => ofP(prod, pri).length);
    const r = base("leakage", p);
    r.detail = { quarter: q.label, counts: c, inProd: !!inProd, list: prod };
    if (!inProd && !prod.length) {
      r.autoRag = r.rag = "Green"; r.notInProd = true;
      r.position = "Not in Production";
      r.metric = { value: 0, text: "0", label: "Production defects", better: "down", fmt: "int" };
      r.highlights.push("The solution has not yet been deployed to production, so there are no production defects to report.");
      r.exec = "Not yet in production — no defect leakage data exists for the quarter.";
      return r;
    }
    const [p1, p2, p3] = c;
    r.autoRag = r.rag = p1 >= 1 || p2 >= 5 || p3 >= 6 ? "Red" : (p2 >= 1 && p2 <= 4) || (p3 >= 1 && p3 <= 5) ? "Amber" : "Green";
    r.position = `${prod.length} in production (${q.label})`;
    r.metric = { value: prod.length, text: String(prod.length), label: "Production defects", better: "down", fmt: "int" };
    r.highlights.push(`${plural(prod.length, "defect")} found in production in ${q.label}: P1 ${p1}, P2 ${p2}, P3 ${p3}.`);
    if (r.rag !== "Green") r.concern.push(`Leakage thresholds breached (Red: P1 ≥ 1, P2 ≥ 5 or P3 ≥ 6; Amber: any P2/P3).`);
    r.exec = `${prod.length} production defect(s) in ${q.label} → ${r.rag}.`;
    return r;
  }

  function kAutomation(data, p, ctx) {
    const s = snapAt(data, p.end);
    if (!s || (!s.tdTotal && !s.tcTotal)) return empty("automation", "No automation data");
    const th = data.th, r = base("automation", p);
    const ps = snapAt(data, ctx.prevEnd);
    const td = s.tdTotal > 0 ? s.tdDone / s.tdTotal : null, tc = s.tcTotal > 0 ? s.tcDone / s.tcTotal : null;
    const ptd = ps && ps.tdTotal > 0 ? ps.tdDone / ps.tdTotal : null, ptc = ps && ps.tcTotal > 0 ? ps.tcDone / ps.tcTotal : null;
    const main = td != null ? td : tc;
    const pmain = td != null ? ptd : ptc;
    const improved = main != null && pmain != null && main > pmain;
    r.autoRag = r.rag = main != null && (main >= th.autoTarget || improved) ? "Green" : "Amber";
    r.position = `${pctS(td)} test data / ${pctS(tc)} test cases`;
    r.metric = { value: main, text: pctS(main), label: "Test data automation", better: "up", fmt: "pct" };
    r.detail = { td, tc, tdTotal: s.tdTotal, tdDone: s.tdDone, tcTotal: s.tcTotal, tcDone: s.tcDone, ptd, ptc, improved, target: th.autoTarget,
      series: data.snapshots.filter((x) => x.weekEnding <= p.end && (x.tdTotal || x.tcTotal)).map((x) => ({ we: x.weekEnding, td: x.tdTotal > 0 ? x.tdDone / x.tdTotal : null, tc: x.tcTotal > 0 ? x.tcDone / x.tcTotal : null })) };
    r.highlights.push(`Test data setup automation covers ${N.int(s.tdDone)} of ${N.int(s.tdTotal)} automatable scenarios (${pct1S(td)}); test-case execution automation covers ${N.int(s.tcDone)} of ${N.int(s.tcTotal)} (${pct1S(tc)}).`);
    if (improved) r.highlights.push(`Test data automation has improved by ${((main - pmain) * 100).toFixed(1)} points since the previous period.`);
    if (main != null && main < th.autoTarget) r.concern.push(`Automation is ${((th.autoTarget - main) * 100).toFixed(0)} points short of the ${pctS(th.autoTarget)} target; ${pctS(1 - main)} of test data preparation remains manual.`);
    r.exec = `Test data automation stands at ${pct1S(td)}${improved ? " and is improving" : ""}. Increasing automation is the main lever to cut preparation effort, accelerate test cycles and improve repeatability.`;
    return r;
  }

  function msEval(m, asOf, th) {
    if (m.completed && m.completed <= asOf) {
      const late = m.completed > m.due;
      return { m, status: late ? "Completed Late" : "Completed", rag: late ? "Red" : "Green", delay: late ? D.diff(m.due, m.completed) : 0 };
    }
    if (asOf > m.due) return { m, status: "Overdue", rag: "Red", delay: D.diff(m.due, asOf) };
    const dd = D.diff(asOf, m.due);
    return { m, status: "On Schedule", rag: dd <= th.msWindow ? "Amber" : "Green", delay: 0, daysTo: dd };
  }
  K.msEval = msEval;

  function kMilestones(data, p, ctx) {
    if (!data.milestones.length) return empty("milestones", "No milestones");
    const th = data.th, asOf = ctx.asOf, r = base("milestones", p);
    const ev = data.milestones.map((m) => msEval(m, asOf, th)).sort((a, b) => (a.m.due < b.m.due ? -1 : 1));
    const c = (s) => ev.filter((e) => e.status === s).length;
    const total = ev.length, completed = c("Completed"), late = c("Completed Late"), overdue = c("Overdue"), onSched = c("On Schedule");
    const amber = ev.filter((e) => e.rag === "Amber").length;
    const adherence = (total - overdue - late) / total;
    r.autoRag = r.rag = ev.some((e) => e.rag === "Red") ? "Red" : amber / total > th.msAmberShare ? "Amber" : "Green";
    const critDelay = Math.max(0, ...ev.filter((e) => e.m.critical).map((e) => e.delay));
    r.position = `${Math.round(adherence * 100)}% on schedule · ${overdue} overdue`;
    r.metric = { value: adherence, text: pctS(adherence), label: "Milestone adherence", better: "up", fmt: "pct" };
    r.detail = { ev, total, completed, late, overdue, onSched, amber, adherence, critDelay };
    r.highlights.push(`${completed + late} of ${total} milestones completed${late ? ` (${late} late)` : ", none completed late"}; ${onSched} on schedule; ${overdue} overdue.`);
    const next = ev.filter((e) => e.status === "On Schedule").slice(0, 3);
    if (next.length) r.highlights.push(`Next due: ${next.map((e) => `${e.m.name} (${D.fmt(e.m.due, false)})`).join("; ")}.`);
    ev.filter((e) => e.status === "Overdue").forEach((e) => r.concern.push(`${e.m.name} is ${plural(e.delay, "day")} overdue against the planned date of ${D.fmt(e.m.due)}${e.m.forecast ? `; recovery forecast ${D.fmt(e.m.forecast)}` : ""}.`));
    r.exec = overdue
      ? `${Math.round(adherence * 100)}% milestone adherence. Schedule variance is concentrated in ${overdue === 1 ? "a single milestone" : overdue + " milestones"} rather than being systemic; focused recovery attention is required.`
      : `All milestones are completed or progressing to plan (${Math.round(adherence * 100)}% adherence).`;
    return r;
  }

  function kEnvironment(data, p, ctx) {
    const snaps = snapsIn(data, p.start, p.end).filter((s) => s.testDays != null);
    if (!snaps.length) return empty("environment", "No environment data");
    const th = data.th, r = base("environment", p);
    const days = snaps.reduce((a, s) => a + (s.testDays || 0), 0), down = snaps.reduce((a, s) => a + (s.downtime || 0), 0);
    const req = days * th.hoursPerTestDay, avail = req > 0 ? 1 - down / req : null;
    const g = p.type === "week" ? th.dtWGreen : th.dtMGreen, a = p.type === "week" ? th.dtWAmber : th.dtMAmber;
    r.autoRag = r.rag = down <= g ? "Green" : down <= a ? "Amber" : "Red";
    r.position = `${pct1S(avail)}`;
    r.metric = { value: avail, text: pct1S(avail), label: "Availability", better: "up", fmt: "pct1" };
    const series = data.snapshots.filter((x) => x.weekEnding <= p.end && x.testDays != null).map((x) => ({ we: x.weekEnding, down: x.downtime || 0, avail: x.testDays > 0 ? 1 - (x.downtime || 0) / (x.testDays * th.hoursPerTestDay) : null }));
    r.detail = { days, down, req, avail, greenAt: g, amberAt: a, series, weeks: snaps };
    r.highlights.push(`${N.round1(down)} hours of test environment downtime across ${days} planned test days (${N.round1(req)} required hours) — ${pct1S(avail)} availability.`);
    r.highlights.push(`Downtime thresholds for the ${p.type}: Green ≤ ${g}h, Amber ≤ ${a}h.`);
    if (r.rag !== "Green") r.concern.push(`Downtime of ${N.round1(down)}h exceeds the ${g}h Green threshold.`);
    r.exec = `Test environment availability was ${pct1S(avail)}, with ${N.round1(down)}h of downtime — ${r.rag === "Green" ? "within" : "outside"} the agreed tolerance and ${r.rag === "Green" ? "supporting delivery of planned testing" : "putting test execution at risk"}.`;
    return r;
  }

  function kTsr(data, p, ctx) {
    const hols = holSet(data), th = data.th, asOf = ctx.asOf;
    const rel = data.tsrs.filter((t) => t.status !== "Withdrawn" && t.received <= asOf && (!t.returned || t.returned > asOf || (t.returned >= p.start && t.returned <= asOf)));
    const items = rel.map((t) => {
      const open = !t.returned || t.returned > asOf;
      const wd = Math.max(0, D.networkdays(t.received, open ? asOf : t.returned, hols) - (t.hold || 0));
      return { t, open, wd, over: wd > th.tsrSla };
    });
    const r = base("tsr", p);
    const recvIn = data.tsrs.filter((t) => t.received >= p.start && t.received <= asOf).length;
    if (!items.length) {
      r.autoRag = r.rag = "Green"; r.position = "No TSRs in period";
      r.detail = { items: [], recvIn, sla: th.tsrSla };
      r.highlights.push("No TSRs required impact assessment during the period."); r.exec = "No service impact attributable to ATS delivery activity.";
      return r;
    }
    const over = items.filter((x) => x.over).length, avg = items.reduce((a, x) => a + x.wd, 0) / items.length;
    r.autoRag = r.rag = over ? "Red" : "Green";
    r.position = over ? `${over} of ${items.length} over ${th.tsrSla}-day SLA` : `${plural(items.length, "TSR")} within SLA`;
    r.metric = { value: avg, text: avg.toFixed(1) + "d", label: "Avg working days", better: "down", fmt: "days" };
    r.detail = { items, over, avg, recvIn, sla: th.tsrSla };
    r.highlights.push(`${plural(items.length, "TSR")} assessed or in progress (${recvIn} received in the period); average ${avg.toFixed(1)} working days against a ${th.tsrSla}-day limit.`);
    items.forEach((x) => r.highlights.push(`${x.t.ref}: ${x.wd} working day${x.wd === 1 ? "" : "s"}${x.open ? " so far (open)" : ""}${x.t.hold ? `, ${x.t.hold} hold day${x.t.hold === 1 ? "" : "s"} excluded` : ""}.`));
    if (over) r.concern.push(`${plural(over, "TSR")} exceeded the ${th.tsrSla} working-day limit: ${items.filter((x) => x.over).map((x) => `${x.t.ref} (${x.wd}d)`).join(", ")}.`);
    r.exec = over ? `${over} TSR(s) breached the ${th.tsrSla}-day assessment limit and need review.` : `All TSRs were assessed within the ${th.tsrSla}-day limit — no service impact attributable to ATS.`;
    return r;
  }

  function kCsat(data, p, ctx) {
    const scored = data.snapshots.filter((s) => s.csat != null && s.weekEnding <= p.end);
    const r = base("csat", p);
    if (!scored.length) {
      r.autoRag = r.rag = "Grey"; r.position = "Not requested"; r.empty = true; r.detail = { scores: [] };
      r.highlights.push("No customer satisfaction score has been requested yet."); r.exec = "No feedback received yet.";
      return r;
    }
    const cur = scored[scored.length - 1], prev = scored.length > 1 ? scored[scored.length - 2] : null, pass = data.th.csatPass;
    const c = cur.csat, pv = prev ? prev.csat : null;
    r.autoRag = r.rag = pv == null ? (c >= pass ? "Green" : "Amber")
      : (c > pv && c >= pass) || (c === pv && c >= pass) ? "Green"
      : (c < pv && c < pass) || (c === pv && c < pass) ? "Red" : "Amber";
    r.position = `${c.toFixed(1)} / 5` + (pv != null ? ` (prev ${pv.toFixed(1)})` : "");
    r.metric = { value: c, text: c.toFixed(1), label: "CSAT score", better: "up", fmt: "num1" };
    r.detail = { cur, prev, pass, scores: scored.map((s) => ({ we: s.weekEnding, score: s.csat, n: s.csatResp })) };
    r.highlights.push(`Current CSAT score ${c.toFixed(1)}/5${cur.csatResp ? ` from ${cur.csatResp} responses` : ""}${pv != null ? `, previous ${pv.toFixed(1)}` : ""}; pass mark ${pass}.`);
    r.exec = `Customer satisfaction is ${c.toFixed(1)}/5 — ${c >= pass ? "above" : "below"} the pass mark${pv != null ? " and " + (c > pv ? "improving" : c < pv ? "declining" : "stable") : ""}.`;
    return r;
  }

  function kResource(data, p, ctx) {
    const s = snapAt(data, p.end);
    if (!s || s.resPlan == null) return empty("resource", "No resourcing data");
    const th = data.th, r = base("resource", p);
    const ratio = s.resPlan > 0 ? s.resAct / s.resPlan : null;
    r.autoRag = r.rag = ratio == null ? "Grey" : ratio >= 1 ? "Green" : ratio >= th.resAmber ? "Amber" : "Red";
    r.position = `${s.resAct} of ${s.resPlan} FTE in place`;
    r.metric = { value: ratio, text: pctS(ratio), label: "Resourced vs plan", better: "up", fmt: "pct" };
    const roster = data.resourcing.filter((x) => x.status !== "Left" && (!x.start || x.start <= p.end) && (!x.end || x.end >= p.end) && x.status !== "Planned");
    const orgs = {}; roster.forEach((x) => { orgs[x.org || "Unassigned"] = (orgs[x.org || "Unassigned"] || 0) + (x.alloc == null ? 1 : x.alloc); });
    r.detail = { plan: s.resPlan, act: s.resAct, ratio, roster, orgs, series: data.snapshots.filter((x) => x.weekEnding <= p.end && x.resPlan != null).map((x) => ({ we: x.weekEnding, plan: x.resPlan, act: x.resAct })), planned: data.resourcing.filter((x) => x.status === "Planned") };
    r.highlights.push(`${s.resAct} of ${s.resPlan} planned FTE are in place (${pctS(ratio)}).`);
    if (roster.length) r.highlights.push(`Active team: ${roster.length} roles${data.config.ptm ? `, led by ${data.config.ptm} (Programme Test Manager)` : ""}${data.config.tm ? ` and ${data.config.tm} (Test Manager)` : ""}.`);
    if (ratio != null && ratio < 1) r.concern.push(`${N.round1(s.resPlan - s.resAct)} FTE gap against plan.`);
    r.exec = ratio != null && ratio >= 1 ? "The team is fully resourced against plan, with clear governance and ownership across SIT execution, defect management and stakeholder engagement." : "Resourcing is below plan; backfill is required to protect delivery.";
    return r;
  }

  function kDemand(data, p, ctx) {
    const m0 = D.monthStart(D.monthKey(p.end));
    const rows = data.demand.filter((d) => d.month >= m0 && d.planned != null).slice(0, 6);
    if (!rows.length) return empty("demand", "No forecast data");
    const r = base("demand", p);
    const next3 = rows.slice(0, 3);
    const gap = (d) => (d.planned || 0) - (d.available == null ? d.planned : d.available);
    const maxGap = Math.max(0, ...next3.map(gap));
    r.autoRag = r.rag = maxGap > 1 ? "Red" : maxGap > 0 ? "Amber" : "Green";
    const peak = rows.reduce((a, d) => (d.planned > a.planned ? d : a), rows[0]);
    r.position = maxGap > 0 ? `Gap of ${N.round1(maxGap)} FTE within 3 months` : "Capacity aligned";
    r.metric = { value: maxGap, text: N.round1(maxGap) + " FTE", label: "Max gap (next 3 months)", better: "down", fmt: "fte" };
    r.detail = { rows: data.demand.filter((d) => d.month >= D.add(m0, -92)), upcoming: rows, maxGap, peak };
    r.highlights.push(`Resource visibility through ${D.fmtMonth(rows[rows.length - 1].month.slice(0, 7))}; peak planned demand ${peak.planned} FTE in ${D.fmtMonth(peak.month.slice(0, 7), false)}.`);
    rows.filter((d) => gap(d) > 0).slice(0, 3).forEach((d) => r.concern.push(`${D.fmtMonth(d.month.slice(0, 7), false)}: ${d.planned} FTE needed vs ${d.available} available (gap ${N.round1(gap(d))}).`));
    r.exec = maxGap > 0 ? "Demand is forecast to exceed available capacity; early recruitment/onboarding decisions are needed to avoid slippage." : "Capacity is aligned to the forecast demand pipeline.";
    return r;
  }

  function kCommercial(data, p, ctx) {
    const s = snapAt(data, p.end), budget = data.config.budget;
    if (!s || (s.spend == null && s.forecast == null)) return empty("commercial", "No commercial data");
    const th = data.th, r = base("commercial", p);
    const fc = s.forecast != null ? s.forecast : null;
    let rag = "Grey";
    if (fc != null && budget) rag = fc <= budget ? "Green" : fc <= budget * (1 + th.comAmber) ? "Amber" : "Red";
    r.autoRag = r.rag = rag;
    const spendPct = budget && s.spend != null ? s.spend / budget : null;
    r.position = fc != null && budget ? (fc <= budget ? "On track to budget" : `Forecast ${pct1S(fc / budget - 1)} over budget`) : "Spend recorded";
    r.metric = { value: fc != null && budget ? fc - budget : null, text: fc != null && budget ? N.gbpK(fc - budget) : "—", label: "Forecast vs budget", better: "down", fmt: "gbp" };
    r.detail = { spend: s.spend, forecast: fc, budget, spendPct, series: data.snapshots.filter((x) => x.weekEnding <= p.end && x.spend != null).map((x) => ({ we: x.weekEnding, spend: x.spend, forecast: x.forecast })) };
    r.highlights.push(`${N.gbp(s.spend)} spent to date${spendPct != null ? ` (${pctS(spendPct)} of the ${N.gbp(budget)} budget)` : ""}; forecast at completion ${N.gbp(fc)}.`);
    if (rag === "Amber" || rag === "Red") r.concern.push(`Forecast exceeds budget by ${N.gbp(fc - budget)} (${pct1S(fc / budget - 1)}).`);
    r.exec = rag === "Green" ? "Commercials are on track: forecast and delivery commitments remain aligned to budget." : "The forecast at completion is above budget — a commercial review is recommended.";
    return r;
  }

  function kExecution(data, p, ctx) {
    const rows = latestTopicRows(data, p.end);
    if (!rows.length) return empty("execution", "No execution data");
    const th = data.th, r = base("execution", p);
    const sum = (arr) => ({ planned: arr.reduce((a, x) => a + x.planned, 0), designed: arr.reduce((a, x) => a + x.designed, 0), passed: arr.reduce((a, x) => a + x.passed, 0), failed: arr.reduce((a, x) => a + x.failed, 0), blocked: arr.reduce((a, x) => a + x.blocked, 0) });
    const t = sum(rows); t.executed = t.passed + t.failed;
    const execPct = t.planned ? t.executed / t.planned : null, passRate = t.executed ? t.passed / t.executed : null, blockedShare = t.planned ? t.blocked / t.planned : 0;
    r.autoRag = r.rag = t.executed === 0 && t.blocked === 0 ? "Grey" : blockedShare <= th.exAmber ? "Green" : blockedShare <= th.exRed ? "Amber" : "Red";
    r.position = `${N.int(t.executed)} of ${N.int(t.planned)} executed (${pctS(execPct)})`;
    r.metric = { value: execPct, text: pctS(execPct), label: "Executed", better: "up", fmt: "pct" };
    const weeks = Array.from(new Set(data.topicProgress.filter((x) => x.weekEnding <= p.end).map((x) => x.weekEnding))).sort();
    const series = weeks.map((we) => { const s2 = sum(latestTopicRows(data, we)); s2.executed = s2.passed + s2.failed; s2.we = we; return s2; });
    const prevRows = latestTopicRows(data, ctx.prevEnd); const pt = sum(prevRows); pt.executed = pt.passed + pt.failed;
    r.detail = { total: t, prevTotal: pt, execPct, passRate, blockedShare, rows: rows.sort((a, b) => b.planned - a.planned), series };
    r.highlights.push(`${N.int(t.executed)} of ${N.int(t.planned)} test cases executed (${pctS(execPct)}): ${N.int(t.passed)} passed, ${N.int(t.failed)} failed${passRate != null ? ` (${pctS(passRate)} pass rate)` : ""}; ${N.int(t.blocked)} blocked.`);
    const d = t.executed - pt.executed; if (d > 0) r.highlights.push(`${N.int(d)} additional test cases executed since the previous period.`);
    if (t.blocked > 0) r.concern.push(`${N.int(t.blocked)} test cases (${pctS(blockedShare)} of plan) are blocked${data.defects.filter((x) => !x.resolved).length ? " by open defects and dependencies" : ""}.`);
    r.exec = `Execution is ${pctS(execPct)} complete with a ${passRate != null ? pctS(passRate) : "—"} pass rate; ${t.blocked ? `${N.int(t.blocked)} blocked test cases are the main drag on progress.` : "no blockers are impacting execution."}`;
    return r;
  }

  function kReadiness(data, p, ctx) {
    const latest = {};
    data.readiness.forEach((x) => { if (x.weekEnding <= p.end && (!latest[x.topic] || x.weekEnding >= latest[x.topic].weekEnding)) latest[x.topic] = x; });
    const rows = Object.values(latest);
    if (!rows.length) return empty("readiness", "No readiness data");
    const r = base("readiness", p);
    r.autoRag = r.rag = R.worst(rows.map((x) => x.overall));
    const nonGreen = rows.filter((x) => x.overall === "Amber" || x.overall === "Red").length;
    r.position = `${rows.length} topics · ${nonGreen} Amber/Red`;
    r.detail = { rows, prev: data.readiness.filter((x) => x.weekEnding < p.start) };
    rows.forEach((x) => r.highlights.push(`${x.topic}: ${x.overall || "n/a"}${x.sitStart ? ` — planned SIT start ${D.fmt(x.sitStart)}` : ""}.`));
    rows.filter((x) => x.roadToGreen).forEach((x) => r.actions.push(`${x.topic}: ${x.roadToGreen}`));
    r.exec = `${nonGreen} of ${rows.length} topics are not yet Green for SIT readiness; the road-to-green actions above are the focus for the coming period.`;
    return r;
  }

  function kRaid(data, p, ctx) {
    const raid = ctx.raid;
    if (!raid || !ATS.raid) return empty("raid", "RAID log not loaded");
    const r = base("raid", p), s = ATS.raid.summary(raid, ctx.asOf, p);
    r.autoRag = r.rag = s.veryHigh ? "Red" : s.high ? "Amber" : "Green";
    r.position = `${s.openRisks} risks · ${s.openIssues} issues · ${s.openDeps} dependencies open`;
    r.metric = { value: s.veryHigh + s.high, text: String(s.veryHigh + s.high), label: "High / Very High items", better: "down", fmt: "int" };
    r.detail = s;
    r.highlights.push(`${s.openRisks} open risks (${s.riskBands.map((b) => b.n + " " + b.label).filter((x, i) => s.riskBands[i].n).join(", ") || "none rated"}); ${s.openIssues} open issues; ${s.openDeps} open dependencies; ${s.unconfirmedAssumptions} unconfirmed assumptions.`);
    if (s.newInPeriod) r.highlights.push(`${s.newInPeriod} RAID item(s) logged and ${s.closedInPeriod} closed/resolved in the period.`);
    if (s.high || s.veryHigh) r.concern.push(`${s.veryHigh + s.high} open risk/issue item(s) rated High or Very High need daily/immediate attention.`);
    r.exec = r.rag === "Green" ? "No High or Very High rated risks or issues are open; the RAID log is under control." : "High-rated risks/issues are open and require active management.";
    return r;
  }

  const FN = { coverage: kCoverage, leakage: kLeakage, dde: kDde, aging: kAging, automation: kAutomation, milestones: kMilestones, environment: kEnvironment, tsr: kTsr, csat: kCsat, resource: kResource, demand: kDemand, commercial: kCommercial, execution: kExecution, readiness: kReadiness, raid: kRaid };

  // ============================================================
  // Commentary merge (manual text from Excel beats auto-draft)
  // ============================================================
  function lines(s) { return String(s || "").split(/\r?\n/).map((x) => x.replace(/^[\s•\-–*]+/, "").trim()).filter(Boolean); }
  function manualFor(data, kpiName, p) {
    const key = N.norm(kpiName);
    let best = null;
    data.commentary.forEach((c) => { if (N.norm(c.kpi) === key && c.weekEnding >= p.start && c.weekEnding <= p.end && (!best || c.weekEnding >= best.weekEnding)) best = c; });
    return best;
  }
  K.manualFor = manualFor;

  function applyManual(res, data, p) {
    res.src = { highlights: "auto", concern: "auto", actions: "auto", exec: "auto" };
    const c = manualFor(data, res.name, p);
    if (!c) return res;
    res.manual = c;
    if (c.highlights) { res.highlights = lines(c.highlights); res.src.highlights = "excel"; }
    if (c.concern) { res.concern = lines(c.concern); res.src.concern = "excel"; }
    if (c.actions) { res.actions = lines(c.actions); res.src.actions = "excel"; }
    if (c.exec) { res.exec = c.exec; res.src.exec = "excel"; }
    if (c.ragOverride) { res.ragOverridden = res.rag !== c.ragOverride; res.rag = c.ragOverride; }
    if (c.positionOverride) { res.position = c.positionOverride; res.positionOverridden = true; }
    return res;
  }

  // ============================================================
  // computeAll
  // ============================================================
  K.computeAll = function (data, period, ext) {
    ext = ext || {};
    const today = ext.today || D.todayIso();
    const asOf = D.minIso(period.end, today);
    const prevP = K.prevPeriod(period);
    const ctx = { asOf, prevEnd: prevP.end, raid: ext.raid || null, today };
    const items = [], byKey = {};
    K.KPI_META.forEach((m) => {
      let res;
      try { res = FN[m.key](data, period, ctx); } catch (e) { console.error("KPI failed", m.key, e); res = empty(m.key, "Error computing"); }
      applyManual(res, data, period);
      items.push(res); byKey[m.key] = res;
    });
    const sc = items.filter((x) => !K.metaByKey[x.key].extra);
    const computed = R.worst(sc.map((x) => x.rag));
    const snap = snapAt(data, period.end);
    const declared = snap && snap.rag ? snap.rag : "";
    const ov = manualFor(data, "Overall", period);
    const rag = (ov && ov.ragOverride) || declared || computed;
    const out = { period, prevPeriod: prevP, asOf, items, byKey, scorecard: sc, overall: { rag, computed, declared: (ov && ov.ragOverride) || declared, summary: (ov && ov.exec) || (snap && snap.summary) || "", manual: ov, snapWeek: snap ? snap.weekEnding : null } };
    if (ext.withPrev !== false) {
      const prevOut = K.computeAll(data, prevP, Object.assign({}, ext, { withPrev: false }));
      out.prev = prevOut;
      items.forEach((it) => { it.prev = prevOut.byKey[it.key]; });
    }
    return out;
  };

  // ============================================================
  // Weekly series for trend sparklines
  // ============================================================
  K.weeklySeries = function (data, endWe, n, ext) {
    const weeks = K.weeks(data).filter((w) => w <= endWe).slice(-n);
    return weeks.map((we) => {
      const p = K.weekPeriod(we);
      const res = K.computeAll(data, p, Object.assign({}, ext, { withPrev: false }));
      return { we, res };
    });
  };

  // ============================================================
  // Data-quality checks (accuracy matters — fines!)
  // ============================================================
  K.quality = function (data, today) {
    today = today || D.todayIso();
    const w = [];
    data.defects.forEach((d) => {
      if ((d.status === "Closed" || d.status === "Resolved") && !d.resolved) w.push({ sheet: "Defect_Log", msg: `${d.id} is ${d.status} but has no Resolved Date — aging cannot be measured.` });
      if (d.resolved && d.resolved < d.raised) w.push({ sheet: "Defect_Log", msg: `${d.id}: Resolved Date is before Raised Date.` });
      if (d.resolved && (d.status === "Open" || d.status === "Re-Open" || d.status === "In Progress")) w.push({ sheet: "Defect_Log", msg: `${d.id} has a Resolved Date but status is ${d.status}.` });
    });
    data.tsrs.forEach((t) => {
      if (t.returned && t.returned < t.received) w.push({ sheet: "TSR_Log", msg: `${t.ref}: Returned Date is before Received Date.` });
      if (t.returned && t.status === "Open") w.push({ sheet: "TSR_Log", msg: `${t.ref} has a Returned Date but status is Open.` });
    });
    data.snapshots.forEach((s) => {
      ["high", "med", "low"].forEach((l) => { if (s[l + "Cov"] != null && s[l + "Reqs"] != null && s[l + "Cov"] > s[l + "Reqs"]) w.push({ sheet: "Weekly_Snapshot", msg: `W/E ${s.weekEnding}: ${l} risk covered requirements exceed total requirements.` }); });
      if (s.tdDone != null && s.tdTotal != null && s.tdDone > s.tdTotal) w.push({ sheet: "Weekly_Snapshot", msg: `W/E ${s.weekEnding}: automated test-data setups exceed automatable.` });
      if (D.dow(s.weekEnding) !== 5) w.push({ sheet: "Weekly_Snapshot", msg: `W/E ${s.weekEnding} is not a Friday.` });
    });
    const last = data.snapshots[data.snapshots.length - 1];
    if (last) { const age = D.diff(last.weekEnding, today); if (age > 9) w.push({ sheet: "Weekly_Snapshot", msg: `Latest weekly snapshot is ${age} days old (W/E ${last.weekEnding}) — add this week's row.` }); }
    data.milestones.forEach((m) => { if (m.completed && m.completed < D.add(m.due, -365)) w.push({ sheet: "Milestones", msg: `${m.name}: completed date looks wrong.` }); });
    return w;
  };

  // ============================================================
  // Narrative helpers shared by pages
  // ============================================================
  K.delta = function (res) {
    if (!res.prev || !res.metric || !res.prev.metric || res.metric.value == null || res.prev.metric.value == null) return null;
    const diff = res.metric.value - res.prev.metric.value;
    const f = res.metric.fmt;
    const mag = f === "pct" || f === "pct1" ? Math.abs(diff * 100).toFixed(1) + " pts" : f === "gbp" ? N.gbpK(Math.abs(diff)) : f === "days" ? Math.abs(diff).toFixed(1) + "d" : f === "num1" ? Math.abs(diff).toFixed(1) : f === "fte" ? Math.abs(diff).toFixed(1) + " FTE" : String(Math.abs(Math.round(diff * 10) / 10));
    if (Math.abs(diff) < 1e-9) return { dir: 0, text: "no change", good: null };
    const up = diff > 0;
    const better = res.metric.better;
    const good = better == null ? null : better === "up" ? up : !up;
    return { dir: up ? 1 : -1, text: (up ? "▲ " : "▼ ") + mag, good };
  };
})();
