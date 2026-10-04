/* ============================================================
   RAID log — parser (reads the existing "ATS - Test RAID Log.xlsx" format as-is)
   + shared render components (heat-map, tables, band bars)
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date, N = ATS.num, R = ATS.rag, X = ATS.xl, esc = ATS.dom.esc;
  const RA = (ATS.raid = ATS.raid || {});

  RA.BANDS = [
    { label: "Very Low", min: 1, max: 4, color: "#9ad8b0" },
    { label: "Low", min: 5, max: 9, color: "#c9e58f" },
    { label: "Medium", min: 10, max: 14, color: "#ffd66b" },
    { label: "High", min: 15, max: 19, color: "#ff9d5c" },
    { label: "Very High", min: 20, max: 99, color: "#ef5b5b" },
  ];
  RA.bandOf = (score) => {
    if (score == null) return null;
    const b = RA.BANDS.find((x) => score >= x.min && score <= x.max);
    return b ? b.label : score < 1 ? "Very Low" : "Very High";
  };
  const bandColor = (label) => (RA.BANDS.find((b) => b.label === label) || { color: "#d0d5dd" }).color;
  RA.bandColor = bandColor;
  const bandRag = (label) => (label === "Very High" ? "Red" : label === "High" ? "Red" : label === "Medium" ? "Amber" : label ? "Green" : "Grey");

  const LIK = { "very low": 1, rare: 1, low: 2, unlikely: 2, possible: 3, medium: 3, likely: 4, "very likely": 5, "almost certain": 5, "very high": 5, high: 4 };
  const IMP = { negligible: 1, "very low": 1, minor: 2, low: 2, moderate: 3, medium: 3, major: 4, high: 4, severe: 5, critical: 5, "very high": 5 };
  const LVL = { "very low": 1, low: 2, medium: 3, high: 4, "very high": 5 };

  RA.detect = (wb) => wb.SheetNames.includes("Risks") && wb.SheetNames.includes("Issues") && (wb.SheetNames.includes("Dependencies") || wb.SheetNames.includes("Assumptions"));

  function scanLabel(aoa, label) {
    const want = N.norm(label);
    for (const row of aoa) for (let c = 0; c < row.length; c++) {
      if (N.norm(row[c]) === want) { for (let k = c + 1; k < row.length; k++) if (row[k] != null && row[k] !== "") return row[k]; }
    }
    return null;
  }
  const real = (r, ...keys) => !!N.str(r.id) && keys.some((k) => N.str(X.field(r, k)));
  const isYes = (v) => /^(y|yes|true|1)/i.test(N.str(v));

  RA.parse = function (wb) {
    if (!RA.detect(wb)) return null;
    const info = {};
    if (wb.Sheets["Summary Sheet"]) {
      const aoa = X.aoa(wb.Sheets["Summary Sheet"]);
      info.project = N.str(scanLabel(aoa, "Project Name")); info.tsr = N.str(scanLabel(aoa, "TSR ID"));
      info.ptm = N.str(scanLabel(aoa, "Programme Test Manager")); info.start = D.toIso(scanLabel(aoa, "Start Date")); info.end = D.toIso(scanLabel(aoa, "End Date"));
      info.version = N.str(scanLabel(aoa, "Version"));
    }
    const risks = X.table(wb, "Risks", ["ID", "Summary Title", "Likelihood", "Impact", "Status"]).filter((r) => real(r, "summarytitle", "description")).map((r) => {
      const l = N.num(X.field(r, "likelihoodrating")) || LIK[N.str(X.field(r, "likelihood")).toLowerCase()] || null;
      const i = N.num(X.field(r, "impactrating")) || IMP[N.str(X.field(r, "impact")).toLowerCase()] || null;
      const score = l && i ? l * i : null;
      return {
        id: N.str(r.id), title: N.str(X.field(r, "summarytitle")), desc: N.str(X.field(r, "description")), category: N.str(X.field(r, "category")),
        likelihood: N.str(X.field(r, "likelihood")), l, impact: N.str(X.field(r, "impact")), i, score, rating: RA.bandOf(score) || N.str(X.field(r, "riskscore")),
        status: N.str(X.field(r, "status")) || "Open", mitigation: N.str(X.field(r, "mitigationdetails", "mitigation")), owner: N.str(X.field(r, "owner")),
        review: D.toIso(X.field(r, "reviewdate")), where: N.str(X.field(r, "whereraised")), prog: isYes(X.field(r, "prograid")), archived: isYes(r.archived),
      };
    });
    const issues = X.table(wb, "Issues", ["ID", "Summary Title", "Full Description", "Priority", "Severity", "Status"]).filter((r) => real(r, "summarytitle", "fulldescription")).map((r) => {
      const p = N.num(X.field(r, "priorityrating")) || LVL[N.str(X.field(r, "priority")).toLowerCase()] || null;
      const s = N.num(X.field(r, "severityrating")) || IMP[N.str(X.field(r, "severity")).toLowerCase()] || null;
      const score = p && s ? p * s : null;
      return {
        id: N.str(r.id), title: N.str(X.field(r, "summarytitle")), desc: N.str(X.field(r, "fulldescription", "description")), tracking: N.str(X.field(r, "trackingid")),
        reporter: N.str(X.field(r, "reporter")), owner: N.str(X.field(r, "owner")), reported: D.toIso(X.field(r, "datereported")), priority: N.str(X.field(r, "priority")), p,
        severity: N.str(X.field(r, "severity")), s, score, rating: RA.bandOf(score) || N.str(X.field(r, "overallissuerating")), status: N.str(X.field(r, "status")) || "Open",
        originalRisk: N.str(X.field(r, "originalriskid")), target: D.toIso(X.field(r, "targetresolutiondate")), actual: D.toIso(X.field(r, "actualresolutiondate")),
        resolution: N.str(X.field(r, "resolutionsummary")), notes: N.str(X.field(r, "notesactions")), archived: isYes(r.archived),
      };
    });
    const assumptions = X.table(wb, "Assumptions", ["ID", "Description", "Raised By", "Confidence Level", "Status"]).filter((r) => real(r, "description")).map((r) => ({
      id: N.str(r.id), desc: N.str(X.field(r, "description")), by: N.str(X.field(r, "raisedby")), logged: D.toIso(X.field(r, "datelogged")), confidence: N.str(X.field(r, "confidencelevel")),
      impact: N.str(X.field(r, "impactifassumption")), validation: N.str(X.field(r, "validationaction")), due: D.toIso(X.field(r, "validationduedate")), status: N.str(X.field(r, "status")) || "Unconfirmed",
      notes: N.str(X.field(r, "notesactions")), archived: isYes(r.archived),
    }));
    const deps = X.table(wb, "Dependencies", ["ID", "Description", "Dependency For (Who Needs This)", "Type", "Priority", "Status"]).filter((r) => real(r, "description")).map((r) => ({
      id: N.str(r.id), desc: N.str(X.field(r, "description")), forWho: N.str(X.field(r, "dependencyfor")), from: N.str(X.field(r, "dependencyfrom")), type: N.str(X.field(r, "type")),
      required: D.toIso(X.field(r, "daterequired")), requestor: N.str(X.field(r, "requestor")), owner: N.str(X.field(r, "owner")), priority: N.str(X.field(r, "priority")),
      p: LVL[N.str(X.field(r, "priority")).toLowerCase()] || null, impact: N.str(X.field(r, "impactifnotmet")), status: N.str(X.field(r, "status")) || "Open", notes: N.str(X.field(r, "notesactions")), archived: isYes(r.archived),
    }));
    // Decisions sheet (optional — older RAID logs have none). Sheet name tolerant: Decisions / Decision Log / Decision.
    const decName = wb.SheetNames.find((n) => /^decisions?( log)?$/i.test(n.trim()));
    const decisions = !decName ? [] : X.table(wb, decName, ["ID", "Summary Title", "Decision", "Status"]).filter((r) => real(r, "summarytitle", "title", "decisiondescription", "description", "decision")).map((r) => ({
      id: N.str(r.id), title: N.str(X.field(r, "summarytitle", "title", "decision")), desc: N.str(X.field(r, "decisiondescription", "description", "decisiondetails")),
      category: N.str(X.field(r, "category", "type")), maker: N.str(X.field(r, "decisionmakerforum", "decisionmaker", "madeby", "forum", "approver")),
      decided: D.toIso(X.field(r, "datedecided", "decisiondate", "date")), rationale: N.str(X.field(r, "rationale", "reason")), impact: N.str(X.field(r, "impactofdecision", "impact")),
      status: N.str(X.field(r, "status")) || "Pending", owner: N.str(X.field(r, "owner")), review: D.toIso(X.field(r, "reviewdate")),
      linked: N.str(X.field(r, "linkedriskissueid", "linkedid", "linkedriskid", "linkeditem")), notes: N.str(X.field(r, "notesactions", "notes")), archived: isYes(r.archived),
    }));
    return { kind: "raid", info, risks, issues, assumptions, deps, decisions };
  };
  RA.describe = (d) => `${d.risks.length} risks · ${d.issues.length} issues · ${d.deps.length} dependencies · ${d.assumptions.length} assumptions · ${(d.decisions || []).length} decisions`;

  // ------------------------------------------------------------------
  // Status helpers + summary
  // ------------------------------------------------------------------
  const lc = (s) => N.str(s).toLowerCase();
  RA.isOpenRisk = (r) => !r.archived && lc(r.status) !== "closed";
  RA.isOpenIssue = (i) => !i.archived && !/^(resolved|closed)/.test(lc(i.status));
  RA.isOpenDep = (d) => !d.archived && lc(d.status) !== "closed";
  // one definition of an assumption's state, shared by the overview counts and the Assumptions page
  RA.assumptionState = (a) => { const k = lc(a.status); return /incorrect/.test(k) ? "Confirmed Incorrect" : /^(closed|withdrawn|superseded|invalid|no longer)/.test(k) ? "Closed" : /^(un(confirmed|validated|verified)|not\b|tbc|to be)/.test(k) ? "Unconfirmed" : /correct|confirmed|validated/.test(k) ? "Confirmed Correct" : "Unconfirmed"; };
  RA.isPendingDecision = (d) => !d.archived && /^(pending|proposed|draft|open|awaiting|tbc)/.test(lc(d.status));

  // data checks for the Data Quality page
  RA.quality = function (raid) {
    const w = [];
    if (!raid) return w;
    [["Risks", raid.risks], ["Issues", raid.issues], ["Assumptions", raid.assumptions], ["Dependencies", raid.deps], ["Decisions", raid.decisions || []]].forEach(([sheet, rows]) => {
      const seen = {}; rows.forEach((r) => { if (seen[r.id]) w.push({ sheet, msg: `ID ${r.id} appears on more than one row.` }); seen[r.id] = 1; });
    });
    raid.risks.filter((r) => !r.archived && RA.isOpenRisk(r) && !r.rating).forEach((r) => w.push({ sheet: "Risks", msg: `${r.id} has no usable Likelihood / Impact, so it is not rated or placed on the heat-map.` }));
    raid.issues.filter((i) => !i.archived && /^(resolved|closed)/i.test(i.status) && !i.actual).forEach((i) => w.push({ sheet: "Issues", msg: `${i.id} is ${i.status} but has no Actual Resolution Date.` }));
    raid.issues.filter((i) => !i.archived && RA.isOpenIssue(i) && !i.rating).forEach((i) => w.push({ sheet: "Issues", msg: `${i.id} has no usable Priority / Severity, so it is not rated.` }));
    (raid.decisions || []).filter((d) => !d.archived && /^(approved|agreed|accepted|rejected)/i.test(d.status) && !d.decided).forEach((d) => w.push({ sheet: "Decisions", msg: `${d.id} is ${d.status} but has no Date Decided.` }));
    return w;
  };

  RA.summary = function (raid, asOf, period) {
    const risks = raid.risks.filter(RA.isOpenRisk), issues = raid.issues.filter(RA.isOpenIssue), deps = raid.deps.filter(RA.isOpenDep);
    const bandCount = (arr) => RA.BANDS.map((b) => ({ label: b.label, color: b.color, n: arr.filter((x) => x.rating === b.label).length }));
    const rb = bandCount(risks), ib = bandCount(issues);
    const cnt = (l) => risks.filter((x) => x.rating === l).length + issues.filter((x) => x.rating === l).length;
    const depP = {}; deps.forEach((d) => { const k = d.priority || "Unrated"; depP[k] = (depP[k] || 0) + 1; });
    const inP = (d) => d && period && d >= period.start && d <= period.end;
    const newIssues = raid.issues.filter((i) => inP(i.reported)).length, newAss = raid.assumptions.filter((a) => inP(a.logged)).length;
    const closed = raid.issues.filter((i) => inP(i.actual)).length;
    const decs = (raid.decisions || []).filter((d) => !d.archived), pending = decs.filter(RA.isPendingDecision);
    const decInP = decs.filter((d) => inP(d.decided));
    const cutoff = D.add(asOf, 30);
    return {
      openRisks: risks.length, openIssues: issues.length, openDeps: deps.length,
      unconfirmedAssumptions: raid.assumptions.filter((a) => !a.archived && RA.assumptionState(a) === "Unconfirmed").length,
      riskBands: rb, issueBands: ib, veryHigh: cnt("Very High"), high: cnt("High"), depByPriority: depP,
      newInPeriod: newIssues + newAss, closedInPeriod: closed, newIssues, newAssumptions: newAss,
      topRisks: risks.slice().sort((a, b) => (b.score || 0) - (a.score || 0)).slice(0, 6),
      topIssues: issues.slice().sort((a, b) => (b.score || 0) - (a.score || 0)).slice(0, 6),
      depsDue: deps.filter((d) => d.required && d.required <= cutoff).sort((a, b) => (a.required < b.required ? -1 : 1)),
      risksOpen: risks, issuesOpen: issues, depsOpen: deps,
      decisionsTotal: decs.length, decisionsPending: pending.length, decisionsInPeriod: decInP.length,
      pendingDecisions: pending.slice().sort((a, b) => ((a.review || "9999") < (b.review || "9999") ? -1 : 1)),
      recentDecisions: decs.filter((d) => !RA.isPendingDecision(d)).sort((a, b) => ((b.decided || "") < (a.decided || "") ? -1 : 1)).slice(0, 8),
      periodDecisions: decInP.sort((a, b) => ((a.decided || "") < (b.decided || "") ? -1 : 1)),
    };
  };

  // ------------------------------------------------------------------
  // Render components
  // ------------------------------------------------------------------
  const trunc = (s, n) => { s = N.str(s); return s.length > n ? s.slice(0, n - 1) + "…" : s; };
  RA.ratingChip = (label) => (label ? `<span class="sx-rate" style="background:${bandColor(label)}22;color:${darker(bandColor(label))};border:1px solid ${bandColor(label)}77">${esc(label)}</span>` : `<span class="chip grey">—</span>`);
  function darker(hex) {
    const n = parseInt(hex.slice(1), 16), r = Math.max(0, (n >> 16) - 90), g = Math.max(0, ((n >> 8) & 255) - 90), b = Math.max(0, (n & 255) - 90);
    return "#" + ((1 << 24) | (r << 16) | (g << 8) | b).toString(16).slice(1);
  }

  // 5x5 likelihood × impact heat-map of open risks. onCell(l,i,list) via data attributes.
  RA.heatmap = function (risks, opts) {
    opts = opts || {};
    const cells = {};
    risks.forEach((r) => { if (r.l && r.i) { const k = r.l + "-" + r.i; (cells[k] = cells[k] || []).push(r); } });
    const likLabels = ["Very Likely", "Likely", "Possible", "Low", "Very Low"], impLabels = ["Negligible", "Minor", "Moderate", "Major", "Severe"];
    let h = `<div class="sx-heat"><div class="sx-heat-y">Likelihood</div><div class="sx-heat-grid">`;
    for (let l = 5; l >= 1; l--) {
      h += `<div class="sx-heat-rowlabel">${likLabels[5 - l]}</div>`;
      for (let i = 1; i <= 5; i++) {
        const list = cells[l + "-" + i] || [], band = RA.bandOf(l * i), col = bandColor(band);
        h += `<button class="sx-heat-cell${list.length ? " has" : ""}" data-l="${l}" data-i="${i}" style="background:${col}${list.length ? "" : "33"}" title="${esc(band)} (${l}×${i}=${l * i})${list.length ? " — " + esc(list.map((x) => x.id).join(", ")) : ""}">${list.length ? `<b>${list.length}</b>` : ""}</button>`;
      }
    }
    h += `<div></div>` + impLabels.map((x) => `<div class="sx-heat-collabel">${x}</div>`).join("");
    h += `</div><div class="sx-heat-x">Impact</div></div>`;
    return h;
  };

  RA.bandBars = function (bands, total) {
    const max = Math.max(1, ...bands.map((b) => b.n));
    return `<div class="sx-bands">${bands.slice().reverse().map((b) => `
      <div class="sx-band"><span class="sx-band-l">${b.label}</span><span class="sx-band-bar"><i style="width:${(b.n / max) * 100}%;background:${b.color}"></i></span><span class="sx-band-n">${b.n}</span></div>`).join("")}</div>`;
  };

  RA.riskTable = function (items, opts) {
    opts = opts || {};
    if (!items.length) return `<div class="sx-empty-note">No open risks.</div>`;
    return `<div class="table-wrap"><table class="data-table sx-raid-table"><thead><tr><th>ID</th><th>Risk</th><th>Rating</th><th>Owner</th><th>Mitigation</th></tr></thead><tbody>${items.map((r) => `
      <tr><td><b>${esc(r.id)}</b></td><td class="sx-wrap">${esc(trunc(r.title, opts.titleLen || 120))}</td><td>${RA.ratingChip(r.rating)}</td>
      <td class="sx-wrap">${esc(trunc(r.owner, 40))}</td><td class="sx-wrap sx-dim">${esc(trunc(r.mitigation, opts.mitLen || 150))}</td></tr>`).join("")}</tbody></table></div>`;
  };
  RA.issueTable = function (items, opts) {
    opts = opts || {};
    if (!items.length) return `<div class="sx-empty-note">No open issues.</div>`;
    return `<div class="table-wrap"><table class="data-table sx-raid-table"><thead><tr><th>ID</th><th>Issue</th><th>Rating</th><th>Status</th><th>Owner</th><th>Target</th></tr></thead><tbody>${items.map((r) => `
      <tr><td><b>${esc(r.id)}</b></td><td class="sx-wrap">${esc(trunc(r.title, opts.titleLen || 120))}</td><td>${RA.ratingChip(r.rating)}</td><td>${esc(r.status)}</td>
      <td class="sx-wrap">${esc(trunc(r.owner, 30))}</td><td>${r.target ? D.fmt(r.target, false) : "—"}</td></tr>`).join("")}</tbody></table></div>`;
  };
  RA.decisionTable = function (items, opts) {
    opts = opts || {};
    if (!items.length) return `<div class="sx-empty-note">No decisions logged.</div>`;
    const chip = (st) => { const k = lc(st); const c = /^approved/.test(k) ? "green" : /^(pending|proposed)/.test(k) ? "amber" : /^rejected/.test(k) ? "red" : "grey"; return `<span class="chip ${c}">${esc(st || "—")}</span>`; };
    return `<div class="table-wrap"><table class="data-table sx-raid-table"><thead><tr><th>ID</th><th>Decision</th><th>Status</th><th>Decided</th><th>By / forum</th><th>Owner</th><th>Linked</th></tr></thead><tbody>${items.map((r) => `
      <tr><td><b>${esc(r.id)}</b></td><td class="sx-wrap">${esc(trunc(r.title, opts.titleLen || 100))}${r.rationale && opts.why ? `<div class="sx-dim">${esc(trunc(r.rationale, 140))}</div>` : ""}</td><td>${chip(r.status)}</td><td>${r.decided ? D.fmt(r.decided, false) : "—"}</td>
      <td class="sx-wrap">${esc(trunc(r.maker, 28))}</td><td class="sx-wrap">${esc(trunc(r.owner, 24))}</td><td>${esc(r.linked || "—")}</td></tr>`).join("")}</tbody></table></div>`;
  };
  RA.depTable = function (items, opts) {
    opts = opts || {};
    if (!items.length) return `<div class="sx-empty-note">No open dependencies.</div>`;
    return `<div class="table-wrap"><table class="data-table sx-raid-table"><thead><tr><th>ID</th><th>Dependency</th><th>Priority</th><th>Needed by</th><th>From</th><th>Owner</th></tr></thead><tbody>${items.map((r) => `
      <tr><td><b>${esc(r.id)}</b></td><td class="sx-wrap">${esc(trunc(r.desc, opts.titleLen || 110))}</td><td>${RA.ratingChip(r.priority)}</td><td>${r.required ? D.fmt(r.required, false) : "TBC"}</td>
      <td class="sx-wrap">${esc(trunc(r.from, 28))}</td><td class="sx-wrap">${esc(trunc(r.owner, 24))}</td></tr>`).join("")}</tbody></table></div>`;
  };
})();
