/* ============================================================
   ATS POAP — Plan on a Page module
   Roadmap (Gantt) | Plan on a Page | Integration Flows (P2P)
   Plain JS, offline, no network. Globals used: XLSX (SheetJS), Chart (unused).
   Public API: window.ATS.poap = { detect, parse, mount, render, ... }
   ============================================================ */
(function () {
  "use strict";
  window.ATS = window.ATS || {};

  // ---------------------------------------------------------------
  // Constants
  // ---------------------------------------------------------------
  var DAY = 864e5;
  var MON = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var MON_FULL = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var STATUSES = ["Not Started", "In Progress", "Complete", "Delayed", "On Hold", "Dates TBC"];
  var RAGS = ["Green", "Amber", "Red"];
  var RAG_HEX = { Green: "#22c55e", Amber: "#f5a524", Red: "#ef4444" };
  var P2P_RES = ["Not Run", "Pass", "Fail", "Blocked", "N/A"];

  var DEFAULT_PILLARS = [
    { name: "Pillar 1", color: "#2563eb" }, { name: "Pillar 2", color: "#0ea5a0" },
    { name: "Pillar 3", color: "#8b5cf6" }, { name: "Penalties P2P (WBS)", color: "#ec4899" },
    { name: "Cross-cutting", color: "#64748b" }
  ];
  var DEFAULT_TYPES = [
    { name: "Test Prep", color: "#8b9df8" }, { name: "Dependencies", color: "#c2a878" },
    { name: "Regression", color: "#0ea5e9" }, { name: "SIT", color: "#3b5bdb" },
    { name: "Smoke / P2P", color: "#06b6d4" }, { name: "Defect Retest", color: "#e86aa6" },
    { name: "TCR / Sign-off", color: "#0f766e" }, { name: "E2E", color: "#7c3aed" },
    { name: "NFR", color: "#c084fc" }, { name: "Automation", color: "#475569" },
    { name: "Holiday / Freeze", color: "#cbd5e1" }, { name: "Other", color: "#94a3b8" }
  ];
  var FALLBACK_COLORS = ["#f97316", "#84cc16", "#14b8a6", "#6366f1", "#d946ef", "#eab308", "#06b6d4", "#f43f5e"];

  // Phase order for the Plan-on-a-Page chevrons (Smoke / P2P is folded into SIT).
  var PHASES = [
    { key: "Test Prep", label: "Test Prep", short: "Prep", types: ["Test Prep"] },
    { key: "Regression", label: "Regression", short: "Regr.", types: ["Regression"] },
    { key: "SIT", label: "SIT", short: "SIT", types: ["SIT", "Smoke / P2P"] },
    { key: "Defect Retest", label: "Defect Retest", short: "Retest", types: ["Defect Retest"] },
    { key: "TCR / Sign-off", label: "TCR / Sign-off", short: "TCR", types: ["TCR / Sign-off"] },
    { key: "E2E", label: "E2E", short: "E2E", types: ["E2E"] },
    { key: "NFR", label: "NFR", short: "NFR", types: ["NFR"] },
    { key: "Automation", label: "Automation", short: "Auto", types: ["Automation"] }
  ];

  // ---------------------------------------------------------------
  // Small utils
  // ---------------------------------------------------------------
  function dnFromYMD(y, m, d) { return Math.floor(Date.UTC(y, m - 1, d) / DAY); }
  function dnToDate(dn) { return new Date(dn * DAY); }
  function dnParts(dn) { var d = new Date(dn * DAY); return { y: d.getUTCFullYear(), m: d.getUTCMonth(), d: d.getUTCDate(), wd: d.getUTCDay() }; }
  function pad2(n) { return (n < 10 ? "0" : "") + n; }
  function dnToISO(dn) { if (dn == null) return ""; var p = dnParts(dn); return p.y + "-" + pad2(p.m + 1) + "-" + pad2(p.d); }
  function fmtD(dn, noYear) {
    if (dn == null) return "—";
    var p = dnParts(dn);
    return pad2(p.d) + " " + MON[p.m] + (noYear ? "" : " " + p.y);
  }
  function fmtShort(dn) { if (dn == null) return ""; var p = dnParts(dn); return p.d + " " + MON[p.m]; }
  function mondayOf(dn) { var wd = dnParts(dn).wd; return dn - ((wd + 6) % 7); }
  function dnFromLocalDate(d) { return dnFromYMD(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function norm(s) { return String(s == null ? "" : s).toLowerCase().replace(/[^a-z0-9%\/ ]+/g, " ").replace(/\s+/g, " ").trim(); }
  function str(v) { return v == null ? "" : String(v).replace(/\s+$/g, "").replace(/^\s+/g, ""); }
  function clamp(v, a, b) { return v < a ? a : (v > b ? b : v); }
  function hexToRgba(hex, a) {
    var h = String(hex || "#94a3b8").replace("#", "");
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    var n = parseInt(h, 16);
    if (isNaN(n)) n = 0x94a3b8;
    return "rgba(" + ((n >> 16) & 255) + "," + ((n >> 8) & 255) + "," + (n & 255) + "," + a + ")";
  }
  function darken(hex, f) {
    var h = String(hex || "#94a3b8").replace("#", "");
    if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    var n = parseInt(h, 16); if (isNaN(n)) n = 0x94a3b8;
    var r = Math.round(((n >> 16) & 255) * f), g = Math.round(((n >> 8) & 255) * f), b = Math.round((n & 255) * f);
    return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
  }
  function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }
  function cls(el, c, on) { if (on) el.classList.add(c); else el.classList.remove(c); }
  function h(tag, attrs, kids) {
    var e = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (!Object.prototype.hasOwnProperty.call(attrs, k) || attrs[k] == null) continue;
      if (k === "class") e.className = attrs[k];
      else if (k === "html") e.innerHTML = attrs[k];
      else if (k === "text") e.textContent = attrs[k];
      else if (k === "style") e.style.cssText = attrs[k];
      else if (k.indexOf("on") === 0 && typeof attrs[k] === "function") e.addEventListener(k.slice(2), attrs[k]);
      else e.setAttribute(k, attrs[k]);
    }
    if (kids) kids.forEach(function (c) { if (c != null) e.appendChild(typeof c === "string" ? document.createTextNode(c) : c); });
    return e;
  }
  function rafThrottle(fn) {
    var pending = false;
    return function () {
      if (pending) return; pending = true;
      var self = this, args = arguments;
      requestAnimationFrame(function () { pending = false; fn.apply(self, args); });
    };
  }

  // ---------------------------------------------------------------
  // Date parsing (-> integer day number or null)
  // ---------------------------------------------------------------
  var TBC_RE = /tbc|tbd|\?\?|to be confirmed|to be decided/i;
  function monthIdx(s) {
    var k = String(s).toLowerCase().slice(0, 3);
    return MON.map(function (m) { return m.toLowerCase(); }).indexOf(k);
  }
  // a real calendar date or nothing (never roll "2026-25-09" or "31/02/2026" over into a different date)
  function ymdCell(y, m, d) {
    if (m < 1 || m > 12 || d < 1 || d > new Date(Date.UTC(y, m, 0)).getUTCDate()) return { dn: null, tbc: false };
    return { dn: dnFromYMD(y, m, d), tbc: false };
  }
  /** returns { dn: number|null, tbc: boolean } */
  function parseDateCell(v) {
    if (v == null || v === "") return { dn: null, tbc: false };
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return { dn: null, tbc: false };
      // SheetJS Dates can be off by minutes/an hour around DST: add 12h then read local parts.
      var d2 = new Date(v.getTime() + 12 * 3600e3);
      return { dn: dnFromYMD(d2.getFullYear(), d2.getMonth() + 1, d2.getDate()), tbc: false };
    }
    if (typeof v === "number") {
      if (v > 20000 && v < 90000) return { dn: Math.floor(v) - 25569, tbc: false };
      return { dn: null, tbc: false };
    }
    var s = String(v).trim();
    if (!s) return { dn: null, tbc: false };
    var m;
    if ((m = /(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return ymdCell(+m[1], +m[2], +m[3]);
    if ((m = /(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/.exec(s))) {
      var dd = +m[1], mm = +m[2], yy = +m[3];
      if (yy < 100) yy += 2000;
      if (mm > 12 && dd <= 12) { var t = dd; dd = mm; mm = t; }
      return ymdCell(yy, mm, dd);
    }
    if ((m = /(\d{1,2})(?:st|nd|rd|th)?[\s\-]+([A-Za-z]{3,9})\.?,?[\s\-]+(\d{2,4})/.exec(s)) && monthIdx(m[2]) >= 0) {
      var y3 = +m[3]; if (y3 < 100) y3 += 2000;
      return ymdCell(y3, monthIdx(m[2]) + 1, +m[1]);
    }
    if ((m = /([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})/.exec(s)) && monthIdx(m[1]) >= 0) {
      return ymdCell(+m[3], monthIdx(m[1]) + 1, +m[2]);
    }
    if (/^\d{5}(\.\d+)?$/.test(s)) return { dn: Math.floor(+s) - 25569, tbc: false };
    if (TBC_RE.test(s)) return { dn: null, tbc: true };
    return { dn: null, tbc: false };
  }
  function dateOut(dn) { return dn == null ? "" : dnToISO(dn); }

  // ---------------------------------------------------------------
  // Parser
  // ---------------------------------------------------------------
  var PLAN_COLS = {
    id: ["id", "bar id"], pillar: ["pillar"], topic: ["topic", "swimlane", "lane"],
    type: ["activity type", "type", "activity"], item: ["item", "bar label", "label", "name", "task"],
    start: ["start", "start date"], end: ["end", "end date", "finish", "finish date"],
    baseStart: ["baseline start", "baseline start date"], baseEnd: ["baseline end", "baseline end date", "baseline finish"],
    pct: ["% complete", "percent complete", "pct complete", "complete", "progress", "% done"],
    status: ["status"], rag: ["rag", "rag status"], owner: ["owner", "lead"],
    dep: ["depends on", "dependencies", "dependency", "predecessors", "depends"],
    scope: ["scope / notes", "scope notes", "scope"], notes: ["notes", "comments", "note"]
  };
  var PLAN_REQ = ["id", "pillar", "topic", "type", "item", "start", "end"];
  var MS_COLS = {
    id: ["id"], pillar: ["pillar"], topic: ["topic"], name: ["milestone", "name", "title"], date: ["date", "milestone date"],
    type: ["type", "milestone type"], status: ["status"], notes: ["notes", "comments"]
  };
  var MS_REQ = ["name", "date", "type"];
  var P2P_COLS = {
    id: ["flow id", "id"], topic: ["topic"], subProcess: ["sub process", "subprocess"], path: ["system path", "path", "flow"],
    testNo: ["test no", "test number", "test"], parent: ["parent test", "parent"], doc: ["document", "doc"],
    date: ["planned date", "date"], status: ["status"], notes: ["notes", "comments"]
  };
  var P2P_REQ = ["id", "path", "status"];

  function findSheet(wb, name) {
    if (!wb || !wb.SheetNames) return null;
    var key = null, lname = name.toLowerCase();
    for (var i = 0; i < wb.SheetNames.length; i++) {
      if (String(wb.SheetNames[i]).trim().toLowerCase() === lname) { key = wb.SheetNames[i]; break; }
    }
    return key == null ? null : wb.Sheets[key];
  }
  function sheetRows(ws) {
    if (!ws) return [];
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, cellDates: true, defval: null, blankrows: true });
  }
  function mapColumns(row, spec, resultKeys) {
    var map = {}, normRow = row.map(norm);
    Object.keys(spec).forEach(function (k) {
      var al = spec[k];
      for (var a = 0; a < al.length; a++) {
        var ix = normRow.indexOf(al[a]);
        if (ix >= 0 && !isUsed(map, ix)) { map[k] = ix; return; }
      }
    });
    return map;
  }
  function isUsed(map, ix) { for (var k in map) if (map[k] === ix) return true; return false; }
  function findHeader(rows, spec, req) {
    for (var r = 0; r < Math.min(rows.length, 14); r++) {
      var row = rows[r]; if (!row) continue;
      var map = mapColumns(row, spec);
      var hit = req.filter(function (k) { return map[k] != null; }).length;
      if (hit >= Math.min(4, req.length)) {
        // result columns for P2P
        return { row: r, map: map, raw: row };
      }
    }
    return null;
  }
  function isExample(notes) { return /^\s*example\b/i.test(String(notes || "")); }

  function canonStatus(s) {
    var n = norm(s);
    if (!n) return "";
    var tbl = { "not started": "Not Started", "notstarted": "Not Started", "in progress": "In Progress", "inprogress": "In Progress", "wip": "In Progress",
      "complete": "Complete", "completed": "Complete", "done": "Complete", "delayed": "Delayed", "late": "Delayed", "on hold": "On Hold", "hold": "On Hold",
      "dates tbc": "Dates TBC", "tbc": "Dates TBC", "tbd": "Dates TBC", "blocked": "Blocked", "planned": "Planned", "achieved": "Achieved", "at risk": "At Risk", "missed": "Missed" };
    return tbl[n] || str(s);
  }
  function canonRag(s) {
    var n = norm(s);
    if (!n) return "";
    if (n === "g" || n === "green") return "Green";
    if (n === "a" || n === "amber" || n === "yellow" || n === "orange") return "Amber";
    if (n === "r" || n === "red") return "Red";
    return "";
  }
  function canonResult(s) {
    var n = norm(s);
    if (!n) return "Not Run";
    if (n === "pass" || n === "passed" || n === "p") return "Pass";
    if (n === "fail" || n === "failed" || n === "f") return "Fail";
    if (n === "blocked") return "Blocked";
    if (n === "n/a" || n === "na" || n === "n a") return "N/A";
    return "Not Run";
  }
  function parsePct(v) {
    if (v == null || v === "") return null;
    var x;
    if (typeof v === "number") x = v;
    else {
      var s = String(v).trim();
      var hasPct = /%/.test(s);
      x = parseFloat(s.replace(/[^0-9.\-]/g, ""));
      if (isNaN(x)) return null;
      if (hasPct) x = x / 100;
    }
    if (isNaN(x)) return null;
    if (x > 1) x = x / 100;
    return clamp(x, 0, 1);
  }
  function splitList(v) {
    return String(v == null ? "" : v).split(/[,;\n]+/).map(function (s) { return s.trim(); }).filter(Boolean);
  }
  function splitPath(p) {
    return String(p == null ? "" : p).split(/\s*(?:→|->|=>|›|>)\s*/).map(function (s) { return s.trim(); }).filter(Boolean);
  }
  function hexClean(v, fb) {
    var s = String(v == null ? "" : v).trim().replace(/^#/, "");
    return /^[0-9a-fA-F]{6}$/.test(s) ? "#" + s.toLowerCase() : (/^[0-9a-fA-F]{3}$/.test(s) ? "#" + s.toLowerCase() : fb);
  }

  function parseConfig(wb) {
    var cfg = { program: "", start: "", end: "", statusDate: "", mode: "", pillars: [], types: [] };
    var ws = findSheet(wb, "POAP_Config");
    var rows = sheetRows(ws);
    var block = null;
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r] || [], a = str(row[0]), b = row[1];
      var na = norm(a);
      if (!na) { if (block && row.every(function (c) { return c == null || c === ""; })) { /* blank row ends block header gap */ } continue; }
      if (na === "pillars") { block = "pillars"; r++; continue; }
      if (na === "activity types") { block = "types"; r++; continue; }
      if (block === "pillars") { cfg.pillars.push({ name: a, color: hexClean(b, "") }); continue; }
      if (block === "types") { cfg.types.push({ name: a, color: hexClean(b, "") }); continue; }
      if (na === "program name") cfg.program = str(b);
      else if (na === "plan start date") cfg.start = dateOut(parseDateCell(b).dn);
      else if (na === "plan end date") cfg.end = dateOut(parseDateCell(b).dn);
      else if (na.indexOf("status date") === 0) cfg.statusDate = dateOut(parseDateCell(b).dn);
      else if (na === "data mode") cfg.mode = str(b).toUpperCase();
    }
    // defaults / colour fill-in
    function fill(list, defs) {
      var out = list.filter(function (x) { return x.name; });
      if (!out.length) out = defs.map(function (d) { return { name: d.name, color: d.color }; });
      out.forEach(function (x, i) {
        if (!x.color) { var d = defs.filter(function (z) { return z.name === x.name; })[0]; x.color = d ? d.color : FALLBACK_COLORS[i % FALLBACK_COLORS.length]; }
      });
      return out;
    }
    cfg.pillars = fill(cfg.pillars, DEFAULT_PILLARS);
    cfg.types = fill(cfg.types, DEFAULT_TYPES);
    return cfg;
  }

  function parsePlan(rows, hd) {
    var out = [], seen = {}, m = hd.map;
    function cell(row, k) { return m[k] == null ? null : row[m[k]]; }
    for (var r = hd.row + 1; r < rows.length; r++) {
      var row = rows[r]; if (!row) continue;
      var item = str(cell(row, "item")), topic = str(cell(row, "topic")), id = str(cell(row, "id")), pillar = str(cell(row, "pillar"));
      var type = str(cell(row, "type"));
      var notes = str(cell(row, "notes"));
      if (isExample(notes)) continue;
      if (!item && !id && !topic && !type) continue;
      if (!item && !topic) continue;
      var sd = parseDateCell(cell(row, "start")), ed = parseDateCell(cell(row, "end"));
      var bs = parseDateCell(cell(row, "baseStart")), be = parseDateCell(cell(row, "baseEnd"));
      var status = canonStatus(cell(row, "status"));
      var rawStatusCells = [cell(row, "start"), cell(row, "end")].join("|");
      var tbcText = sd.tbc || ed.tbc || TBC_RE.test(rawStatusCells);
      if (tbcText && !status) status = "Dates TBC";
      var pct = parsePct(cell(row, "pct"));
      var noId = !id, dupOf = "";
      if (!id) id = "ROW-" + (r + 1);
      if (seen[id]) { dupOf = id; var n = 2; while (seen[id + "~" + n]) n++; id = id + "~" + n; }
      seen[id] = 1;
      out.push({
        noId: noId, dupOf: dupOf, id: id, pillar: pillar, topic: topic, type: type, item: item || topic,
        start: dateOut(sd.dn), end: dateOut(ed.dn), baseStart: dateOut(bs.dn), baseEnd: dateOut(be.dn),
        pct: pct, status: status, rag: canonRag(cell(row, "rag")), owner: str(cell(row, "owner")),
        dependsOn: splitList(cell(row, "dep")), scope: str(cell(row, "scope")), notes: notes, row: r + 1
      });
    }
    return out;
  }

  function parseMilestones(rows, hd) {
    var out = [], seen = {}, m = hd.map;
    function cell(row, k) { return m[k] == null ? null : row[m[k]]; }
    for (var r = hd.row + 1; r < rows.length; r++) {
      var row = rows[r]; if (!row) continue;
      var name = str(cell(row, "name")), notes = str(cell(row, "notes"));
      if (isExample(notes)) continue;
      var dt = parseDateCell(cell(row, "date"));
      if (!name && dt.dn == null) continue;
      if (!name) name = "Milestone";
      var id = str(cell(row, "id")) || ("MS-" + (r + 1));
      if (seen[id]) { var n = 2; while (seen[id + "~" + n]) n++; id = id + "~" + n; }
      seen[id] = 1;
      out.push({ id: id, pillar: str(cell(row, "pillar")), topic: str(cell(row, "topic")), name: name, date: dateOut(dt.dn),
        type: str(cell(row, "type")) || "Other", status: canonStatus(cell(row, "status")) || "Planned", notes: notes, row: r + 1 });
    }
    return out;
  }

  function parseP2P(rows, hd) {
    var out = [], seen = {}, m = hd.map;
    // result columns: header starting with 2xx/3xx/4xx/5xx
    var res = {};
    hd.raw.forEach(function (hh, ix) {
      var n = norm(hh), mm = /^([2345])xx/.exec(n);
      if (mm) res["r" + mm[1] + "xx"] = ix;
    });
    function cell(row, k) { return m[k] == null ? null : row[m[k]]; }
    for (var r = hd.row + 1; r < rows.length; r++) {
      var row = rows[r]; if (!row) continue;
      var path = str(cell(row, "path")), notes = str(cell(row, "notes"));
      if (isExample(notes)) continue;
      var id = str(cell(row, "id"));
      if (!path && !id) continue;
      if (!id) id = "FL-" + (r + 1);
      if (seen[id]) { var n = 2; while (seen[id + "~" + n]) n++; id = id + "~" + n; }
      seen[id] = 1;
      var dt = parseDateCell(cell(row, "date"));
      var status = canonStatus(cell(row, "status"));
      if (!status) status = dt.tbc ? "Dates TBC" : "Not Started";
      var o = {
        id: id, topic: str(cell(row, "topic")), subProcess: str(cell(row, "subProcess")), path: path, pathNodes: splitPath(path),
        testNo: str(cell(row, "testNo")), parent: str(cell(row, "parent")), doc: str(cell(row, "doc")),
        date: dateOut(dt.dn), status: status, notes: notes, row: r + 1
      };
      ["r2xx", "r3xx", "r4xx", "r5xx"].forEach(function (k) { o[k] = canonResult(res[k] == null ? null : row[res[k]]); });
      out.push(o);
    }
    return out;
  }

  function detect(wb) {
    try { return !!findSheet(wb, "POAP_Plan"); } catch (e) { return false; }
  }

  function parse(wb) {
    try {
      var wsPlan = findSheet(wb, "POAP_Plan");
      if (!wsPlan) return null;
      var config = parseConfig(wb);
      var pr = sheetRows(wsPlan), phd = findHeader(pr, PLAN_COLS, PLAN_REQ);
      var plan = phd ? parsePlan(pr, phd) : [];
      var ms = [];
      var wsMs = findSheet(wb, "POAP_Milestones");
      if (wsMs) { var mr = sheetRows(wsMs), mhd = findHeader(mr, MS_COLS, MS_REQ); if (mhd) ms = parseMilestones(mr, mhd); }
      var p2p = [];
      var wsP = findSheet(wb, "P2P_Matrix");
      if (wsP) { var qr = sheetRows(wsP), qhd = findHeader(qr, P2P_COLS, P2P_REQ); if (qhd) p2p = parseP2P(qr, qhd); }
      return { config: config, plan: plan, milestones: ms, p2p: p2p };
    } catch (e) {
      ATSpoap.lastError = e;
      return null;
    }
  }

  // ---------------------------------------------------------------
  // State + model
  // ---------------------------------------------------------------
  var ATSpoap = window.ATS.poap = window.ATS.poap || {};
  var S = {
    root: null, data: null, ctx: {}, model: null,
    view: "roadmap", zoom: "month", pxd: 2.6,
    filters: { pillar: "", type: "", status: "", rag: "", q: "" },
    p2pF: { topic: "", status: "", result: "", sys: [] },
    collapsed: {}, selKind: null, selId: null, hoverId: null,
    dens: "comfortable",
    fb: [], fm: [], fp: [], dirty: { roadmap: true, poap: true, p2p: true },
    els: {}, views: {}
  };

  function phaseOf(type) {
    for (var i = 0; i < PHASES.length; i++) if (PHASES[i].types.indexOf(type) >= 0) return i;
    return -1;
  }

  function buildModel(data, ctx) {
    var cfg = data.config || {};
    var todayDn = (ctx && ctx.today instanceof Date && !isNaN(ctx.today)) ? dnFromLocalDate(ctx.today) : dnFromLocalDate(new Date());
    var sd = cfg.statusDate ? parseDateCell(cfg.statusDate).dn : null;
    var statusDn = sd != null ? sd : todayDn;
    var M = { cfg: cfg, todayDn: todayDn, statusDn: statusDn, program: String(cfg.program || "Plan on a Page").replace(/\s*[\(\[]\s*sample\s*[\)\]]\s*$/i, "") };
    // pillars / types
    var pillars = (cfg.pillars || []).map(function (p) { return { name: p.name, color: p.color }; });
    var types = (cfg.types || []).map(function (p) { return { name: p.name, color: p.color }; });
    function ensure(list, name, fbIdx) {
      if (!name) return;
      if (!list.some(function (x) { return x.name === name; })) list.push({ name: name, color: FALLBACK_COLORS[fbIdx % FALLBACK_COLORS.length] });
    }
    (data.plan || []).forEach(function (p, i) { ensure(pillars, p.pillar, i); ensure(types, p.type, i); });
    (data.milestones || []).forEach(function (p, i) { ensure(pillars, p.pillar, i); });
    M.pillars = pillars; M.types = types;
    M.pillarColor = {}; pillars.forEach(function (p) { M.pillarColor[p.name] = p.color; });
    M.typeColor = {}; types.forEach(function (p) { M.typeColor[p.name] = p.color; });
    M.pillarIdx = {}; pillars.forEach(function (p, i) { M.pillarIdx[p.name] = i; });
    M.typeIdx = {}; types.forEach(function (p, i) { M.typeIdx[p.name] = i; });

    var byId = {};
    var bars = (data.plan || []).map(function (p, i) {
      var s = p.start ? parseDateCell(p.start).dn : null, e = p.end ? parseDateCell(p.end).dn : null;
      var bs = p.baseStart ? parseDateCell(p.baseStart).dn : null, be = p.baseEnd ? parseDateCell(p.baseEnd).dn : null;
      if (s != null && e == null) e = s;
      if (e != null && s == null) s = e;
      if (s != null && e < s) { var t = s; s = e; e = t; }
      if (bs != null && be == null) be = bs;
      if (be != null && bs == null) bs = be;
      var status = p.status;
      var noDates = (s == null);
      var b = {
        kind: "bar", id: p.id, idx: i, pillar: p.pillar || "Unassigned", topic: p.topic || "(No topic)", type: p.type || "Other", item: p.item,
        s: s, e: e, bs: bs, be: be, pct: p.pct, status: status, rag: p.rag, owner: p.owner, deps: (p.dependsOn || []).slice(),
        scope: p.scope || "", notes: p.notes || "", wbRow: p.row, raw: p
      };
      if (!b.status) {
        if (noDates) b.status = "Dates TBC";
        else if (e < statusDn) b.status = "Complete";
        else if (s <= statusDn) b.status = "In Progress";
        else b.status = "Not Started";
      }
      if (b.pct == null) {
        if (b.status === "Complete") b.pct = 1;
        else if (b.status === "In Progress" && !noDates) b.pct = clamp((statusDn - s + 1) / (e - s + 1), 0, 1);
        else b.pct = 0;
      }
      b.tbc = (b.status === "Dates TBC") || noDates;
      b.indicative = noDates;
      b.slip = (be != null && e != null) ? e - be : 0;
      b.isHol = (b.type === "Holiday / Freeze");
      b.color = M.typeColor[b.type] || "#94a3b8";
      b.key = b.pillar + "|" + b.topic;
      byId[b.id] = b;
      return b;
    });
    // indicative positions for bars without dates: after the latest earlier-phase bar in the topic
    var lanes = {};
    bars.forEach(function (b) { (lanes[b.key] = lanes[b.key] || []).push(b); });
    Object.keys(lanes).forEach(function (k) {
      var grp = lanes[k], dated = grp.filter(function (x) { return x.s != null && !x.isHol; });
      var maxEnd = dated.length ? Math.max.apply(null, dated.map(function (x) { return x.e; })) : statusDn;
      grp.forEach(function (b) {
        if (b.s != null) return;
        var ph = phaseOf(b.type);
        var prior = dated.filter(function (x) { var px = phaseOf(x.type); return px >= 0 && ph >= 0 ? px <= ph : true; });
        var base = prior.length ? Math.max.apply(null, prior.map(function (x) { return x.e; })) : maxEnd;
        b.ps = base + 3; b.pe = b.ps + 27; // indicative window (28 days)
      });
    });
    bars.forEach(function (b) { if (b.s != null) { b.ps = b.s; b.pe = b.e; } b.dur = b.s != null ? b.e - b.s + 1 : null; });
    // dependencies
    bars.forEach(function (b) { b.deps = b.deps.filter(function (id) { return byId[id] && id !== b.id; }); b.succ = []; });
    bars.forEach(function (b) { b.deps.forEach(function (id) { byId[id].succ.push(b.id); }); });
    M.bars = bars; M.byId = byId;

    var ms = (data.milestones || []).map(function (m, i) {
      var dn = m.date ? parseDateCell(m.date).dn : null;
      return { kind: "ms", id: m.id, idx: i, pillar: m.pillar || "", topic: m.topic || "", name: m.name, dn: dn, type: m.type, status: m.status || "Planned", notes: m.notes || "", raw: m };
    }).filter(function (m) { return m.dn != null; });
    ms.forEach(function (m) {
      m.key = m.pillar + "|" + m.topic;
      m.color = m.status === "Missed" ? "#ef4444" : (m.status === "At Risk" ? "#f5a524" : (m.status === "Achieved" ? "#22c55e" : "#0f1c3f"));
      m.key2 = (m.type === "SIT Start" || m.type === "SIT End") ? "sit" : "key";
    });
    M.ms = ms; M.msById = {}; ms.forEach(function (m) { M.msById[m.id] = m; });

    var p2p = (data.p2p || []).map(function (p, i) {
      var dn = p.date ? parseDateCell(p.date).dn : null;
      var o = { kind: "flow", idx: i, dn: dn, raw: p };
      for (var k in p) o[k] = p[k];
      o.kind = "flow"; o.dn = dn;
      return o;
    });
    M.p2p = p2p; M.flowById = {}; p2p.forEach(function (f) { M.flowById[f.id] = f; });

    // range
    var all = [];
    bars.forEach(function (b) { if (b.s != null) { all.push(b.s, b.e); if (b.bs != null) all.push(b.bs, b.be); } else { all.push(b.ps, b.pe); } });
    ms.forEach(function (m) { all.push(m.dn); });
    if (cfg.start) { var cs = parseDateCell(cfg.start).dn; if (cs != null) all.push(cs); }
    if (cfg.end) { var ce = parseDateCell(cfg.end).dn; if (ce != null) all.push(ce); }
    all.push(statusDn, todayDn);
    var mn = Math.min.apply(null, all), mx = Math.max.apply(null, all);
    var a = dnParts(mn - 10); var z = dnParts(mx + 21);
    M.minDn = dnFromYMD(a.y, a.m + 1, 1);
    M.maxDn = dnFromYMD(z.y, z.m + 2, 1) - 1;
    M.sample = false;
    return M;
  }

  // ---------------------------------------------------------------
  // Filtering
  // ---------------------------------------------------------------
  function ragOf(b) { return b.rag || ""; }
  function matchQ(q, vals) {
    if (!q) return true;
    var s = vals.join(" \u0001 ").toLowerCase();
    var parts = q.toLowerCase().split(/\s+/).filter(Boolean);
    for (var i = 0; i < parts.length; i++) if (s.indexOf(parts[i]) < 0) return false;
    return true;
  }
  function passBar(b) {
    var F = S.filters;
    if (F.pillar && b.pillar !== F.pillar) return false;
    if (F.type && b.type !== F.type) return false;
    if (F.status && b.status !== F.status) return false;
    if (F.rag && ragOf(b) !== F.rag) return false;
    return matchQ(F.q, [b.id, b.pillar, b.topic, b.type, b.item, b.owner, b.scope, b.notes, b.status]);
  }
  function barFiltersActive() { var F = S.filters; return !!(F.type || F.status || F.rag || F.q); }
  function passMs(m) {
    var F = S.filters;
    if (F.pillar && m.pillar && m.pillar !== F.pillar) return false;
    if (F.q && !matchQ(F.q, [m.id, m.pillar, m.topic, m.name, m.type, m.notes, m.status])) return false;
    return true;
  }
  function passFlow(f) {
    var F = S.filters, P = S.p2pF;
    if (F.q && !matchQ(F.q, [f.id, f.topic, f.subProcess, f.path, f.testNo, f.parent, f.doc, f.notes, f.status])) return false;
    if (P.topic && f.topic !== P.topic) return false;
    if (P.status && f.status !== P.status) return false;
    if (P.result) {
      var rs = [f.r2xx, f.r3xx, f.r4xx, f.r5xx];
      if (rs.indexOf(P.result) < 0) return false;
    }
    if (P.sys.length) {
      var nodes = f.pathNodes || [];
      var all = P.sys.every(function (sy) { return nodes.indexOf(sy) >= 0; });
      if (!all) return false;
    }
    return true;
  }
  function applyFilters() {
    var M = S.model;
    if (!M) { S.fb = []; S.fm = []; S.fp = []; return; }
    S.fb = M.bars.filter(function (b) { return !b.isHol && passBar(b); });
    S.fm = M.ms.filter(passMs);
    S.fp = M.p2p.filter(passFlow);
  }

  // ---------------------------------------------------------------
  // KPI helpers
  // ---------------------------------------------------------------
  function computeKpis() {
    var M = S.model, bars = S.fb;
    var k = { bars: bars.length, total: M.bars.filter(function (b) { return !b.isHol; }).length, prog: 0, done: 0, delayed: 0, tbc: 0, next: null };
    bars.forEach(function (b) {
      if (b.status === "In Progress") k.prog++;
      else if (b.status === "Complete") k.done++;
      if (b.status === "Delayed" || b.slip > 0) k.delayed++;
      if (b.status === "Dates TBC") k.tbc++;
    });
    var up = S.fm.filter(function (m) { return m.dn >= M.statusDn && m.status !== "Achieved"; }).sort(function (a, b) { return a.dn - b.dn; });
    if (up.length) k.next = up[0];
    return k;
  }
  function statusKind(st) {
    return { "Complete": "green", "In Progress": "teal", "Delayed": "red", "Not Started": "grey", "On Hold": "amber", "Dates TBC": "dash",
      "Blocked": "red", "Planned": "grey", "Achieved": "green", "At Risk": "amber", "Missed": "red" }[st] || "grey";
  }
  function chipHtml(text, kind) { return '<span class="poap-chip poap-chip-' + kind + '">' + esc(text) + "</span>"; }
  function ragDot(rag, title) {
    return '<i class="poap-rag poap-rag-' + (rag ? rag.toLowerCase() : "none") + '" title="' + esc(title || (rag ? "RAG: " + rag : "RAG not set")) + '"></i>';
  }
  function slipText(b) { return b.slip > 0 ? "+" + b.slip + "d" : (b.slip < 0 ? b.slip + "d" : "0d"); }

  // ---------------------------------------------------------------
  // Tooltip (shared)
  // ---------------------------------------------------------------
  var tipEl = null;
  function ensureTip() {
    if (tipEl && document.body.contains(tipEl)) return tipEl;
    tipEl = document.getElementById("poap-tip");
    if (!tipEl) { tipEl = h("div", { id: "poap-tip", class: "poap-tip", role: "tooltip" }); document.body.appendChild(tipEl); }
    return tipEl;
  }
  function showTip(html, evt) {
    var t = ensureTip();
    t.innerHTML = html;
    t.classList.add("on");
    moveTip(evt);
  }
  function moveTip(evt) {
    if (!tipEl || !evt) return;
    var w = tipEl.offsetWidth, hgt = tipEl.offsetHeight;
    var x = evt.clientX + 16, y = evt.clientY + 18;
    if (x + w > window.innerWidth - 10) x = evt.clientX - w - 14;
    if (y + hgt > window.innerHeight - 10) y = evt.clientY - hgt - 14;
    tipEl.style.left = Math.max(6, x) + "px"; tipEl.style.top = Math.max(6, y) + "px";
  }
  function hideTip() { if (tipEl) tipEl.classList.remove("on"); }

  function barTipHtml(b) {
    var M = S.model;
    var dates = b.s != null ? fmtD(b.s) + " → " + fmtD(b.e) + " · " + plural(b.dur, "day") : (b.indicative ? "Dates TBC (indicative position)" : "—");
    var h2 = '<div class="poap-tip-h"><i class="poap-sw" style="background:' + b.color + '"></i>' + esc(b.type) + "<span>" + esc(b.id) + "</span></div>" +
      '<div class="poap-tip-t">' + esc(b.item) + "</div>" +
      '<div class="poap-tip-s">' + esc(b.pillar) + " › " + esc(b.topic) + "</div>" +
      '<div class="poap-tip-r">' + esc(dates) + "</div>" +
      '<div class="poap-tip-r">' + chipHtml(b.status, statusKind(b.status)) + " " + ragDot(b.rag) + " " + Math.round(b.pct * 100) + "% complete" +
      (b.slip ? ' <b class="' + (b.slip > 0 ? "poap-neg" : "poap-pos") + '">' + slipText(b) + " vs baseline</b>" : "") + "</div>";
    return h2;
  }
  function msTipHtml(m) {
    return '<div class="poap-tip-h"><i class="poap-dia" style="background:' + m.color + '"></i>Milestone · ' + esc(m.type) + "<span>" + esc(m.id) + "</span></div>" +
      '<div class="poap-tip-t">' + esc(m.name) + "</div>" +
      '<div class="poap-tip-s">' + esc([m.pillar, m.topic].filter(Boolean).join(" › ") || "Programme-wide") + "</div>" +
      '<div class="poap-tip-r">' + fmtD(m.dn) + " · " + relDays(m.dn) + " " + chipHtml(m.status, statusKind(m.status)) + "</div>";
  }
  function relDays(dn) {
    var d = dn - S.model.statusDn;
    if (d === 0) return "today";
    return d > 0 ? "in " + plural(d, "day") : plural(-d, "day") + " ago";
  }

  // ---------------------------------------------------------------
  // Drawer (details)
  // ---------------------------------------------------------------
  var drawerEl = null, drawerBody = null, drawerPrevFocus = null;
  function ensureDrawer() {
    if (drawerEl && document.body.contains(drawerEl)) return drawerEl;
    drawerEl = document.getElementById("poap-drawer");
    if (!drawerEl) {
      drawerEl = h("aside", { id: "poap-drawer", class: "poap-drawer", role: "dialog", "aria-modal": "false", "aria-label": "Details", tabindex: "-1" });
      drawerEl.innerHTML = '<button class="poap-drawer-x" type="button" aria-label="Close details" title="Close (Esc)">&times;</button><div class="poap-drawer-body"></div>';
      document.body.appendChild(drawerEl);
      drawerEl.querySelector(".poap-drawer-x").addEventListener("click", closeDrawer);
      drawerEl.addEventListener("click", function (ev) {
        var t = ev.target.closest("[data-poap-focus]");
        if (t) { ev.preventDefault(); focusBar(t.getAttribute("data-poap-focus")); return; }
        var g = ev.target.closest("[data-poap-open]");
        if (g) { ev.preventDefault(); openDrawer(g.getAttribute("data-poap-kind") || "bar", g.getAttribute("data-poap-open")); return; }
        var s = ev.target.closest("[data-poap-show]");
        if (s) { ev.preventDefault(); showOnRoadmap(s.getAttribute("data-poap-show")); return; }
        var f = ev.target.closest("[data-poap-sys]");
        if (f) { ev.preventDefault(); toggleSystem(f.getAttribute("data-poap-sys")); setView("p2p"); }
      });
      document.addEventListener("keydown", function (ev) {
        if (ev.key === "Escape" && drawerEl.classList.contains("open")) { closeDrawer(); }
      });
    }
    drawerBody = drawerEl.querySelector(".poap-drawer-body");
    return drawerEl;
  }
  function fact(label, val) { return '<div class="poap-fact"><span>' + esc(label) + "</span><b>" + (val == null || val === "" ? "—" : val) + "</b></div>"; }
  function paraHtml(text) {
    if (!text) return "";
    return String(text).split(/\s\|\s|\n+/).map(function (t) { return t.trim(); }).filter(Boolean).map(function (t) { return "<p>" + esc(t) + "</p>"; }).join("");
  }
  function linkBar(id) {
    var o = S.model.byId[id]; if (!o) return "";
    return '<a href="#" class="poap-link" data-poap-focus="' + esc(id) + '"><i class="poap-sw" style="background:' + o.color + '"></i><span class="poap-link-t">' + esc(o.item) + '</span><span class="poap-link-m">' +
      esc(o.id) + " · " + (o.s != null ? fmtD(o.s, true) + "–" + fmtD(o.e, true) : "dates TBC") + "</span></a>";
  }
  function barDrawerHtml(b) {
    var rows = "";
    var slipChip = b.slip ? chipHtml(slipText(b) + " vs baseline", b.slip > 0 ? "red" : "green") : "";
    rows += '<div class="poap-dr-top" style="--c:' + b.color + '"><span class="poap-dr-type"><i class="poap-sw" style="background:' + b.color + '"></i>' + esc(b.type) + "</span><span class=\"poap-dr-id\">" + esc(b.id) + "</span></div>";
    rows += '<h3 class="poap-dr-title">' + esc(b.item) + "</h3>";
    rows += '<div class="poap-dr-path"><span style="--c:' + (S.model.pillarColor[b.pillar] || "#64748b") + '" class="poap-pdot"></span>' + esc(b.pillar) + " › " + esc(b.topic) + "</div>";
    rows += '<div class="poap-dr-chips">' + chipHtml(b.status, statusKind(b.status)) + (b.rag ? chipHtml("RAG " + b.rag, b.rag.toLowerCase()) : chipHtml("RAG not set", "grey")) + slipChip + (b.indicative ? chipHtml("Indicative dates", "dash") : "") + "</div>";
    rows += '<div class="poap-dr-prog"><div class="poap-dr-progh"><span>Progress</span><b>' + Math.round(b.pct * 100) + '%</b></div><div class="poap-meter"><i style="width:' + Math.round(b.pct * 100) + "%;background:" + b.color + '"></i></div></div>';
    var wk = b.dur != null ? (Math.round(b.dur / 7 * 10) / 10) : null;
    rows += '<div class="poap-facts">' +
      fact("Start", b.s != null ? fmtD(b.s) : "TBC") + fact("End", b.s != null ? fmtD(b.e) : "TBC") +
      fact("Duration", b.dur != null ? b.dur + " days (" + wk + " wks)" : "—") + fact("Slip", b.bs != null ? '<span class="' + (b.slip > 0 ? "poap-neg" : (b.slip < 0 ? "poap-pos" : "")) + '">' + slipText(b) + "</span>" : "No baseline") +
      fact("Baseline start", b.bs != null ? fmtD(b.bs) : "—") + fact("Baseline end", b.be != null ? fmtD(b.be) : "—") +
      fact("Owner", esc(b.owner)) + fact("Workbook row", b.wbRow ? "POAP_Plan!" + b.wbRow : "") + "</div>";
    if (b.scope) rows += '<div class="poap-dr-sec"><h4>Scope / notes</h4><div class="poap-prose">' + paraHtml(b.scope) + "</div></div>";
    if (b.notes) rows += '<div class="poap-dr-sec"><h4>Notes</h4><div class="poap-prose">' + paraHtml(b.notes) + "</div></div>";
    if (b.deps.length) rows += '<div class="poap-dr-sec"><h4>Depends on (' + b.deps.length + ')</h4><div class="poap-links">' + b.deps.map(linkBar).join("") + "</div></div>";
    if (b.succ.length) rows += '<div class="poap-dr-sec"><h4>Blocks / feeds (' + b.succ.length + ')</h4><div class="poap-links">' + b.succ.map(linkBar).join("") + "</div></div>";
    var siblings = S.model.ms.filter(function (m) { return m.key === b.key; });
    if (siblings.length) rows += '<div class="poap-dr-sec"><h4>Topic milestones</h4><div class="poap-links">' + siblings.map(function (m) {
      return '<a href="#" class="poap-link" data-poap-open="' + esc(m.id) + '" data-poap-kind="ms"><i class="poap-dia" style="background:' + m.color + '"></i><span class="poap-link-t">' + esc(m.name) + '</span><span class="poap-link-m">' + fmtD(m.dn, true) + "</span></a>";
    }).join("") + "</div></div>";
    rows += '<div class="poap-dr-actions"><button type="button" class="btn btn-primary btn-sm" data-poap-show="' + esc(b.id) + '">Show on roadmap</button></div>';
    return rows;
  }
  function msDrawerHtml(m) {
    var r = '<div class="poap-dr-top"><span class="poap-dr-type"><i class="poap-dia" style="background:' + m.color + '"></i>Milestone · ' + esc(m.type) + '</span><span class="poap-dr-id">' + esc(m.id) + "</span></div>";
    r += '<h3 class="poap-dr-title">' + esc(m.name) + "</h3>";
    r += '<div class="poap-dr-path">' + esc([m.pillar, m.topic].filter(Boolean).join(" › ") || "Programme-wide") + "</div>";
    r += '<div class="poap-dr-chips">' + chipHtml(m.status, statusKind(m.status)) + chipHtml(relDays(m.dn), "grey") + "</div>";
    r += '<div class="poap-facts">' + fact("Date", fmtD(m.dn)) + fact("Type", esc(m.type)) + fact("Days to go", String(m.dn - S.model.statusDn)) + fact("Workbook row", m.raw && m.raw.row ? "POAP_Milestones!" + m.raw.row : "") + "</div>";
    if (m.notes) r += '<div class="poap-dr-sec"><h4>Notes</h4><div class="poap-prose">' + paraHtml(m.notes) + "</div></div>";
    var near = S.model.bars.filter(function (b) { return b.key === m.key && !b.isHol; }).slice(0, 8);
    if (near.length) r += '<div class="poap-dr-sec"><h4>Activities in this topic</h4><div class="poap-links">' + near.map(function (b) { return linkBar(b.id); }).join("") + "</div></div>";
    r += '<div class="poap-dr-actions"><button type="button" class="btn btn-primary btn-sm" data-poap-show="ms:' + esc(m.id) + '">Show on roadmap</button></div>';
    return r;
  }
  function resChip(label, res) {
    var k = { "Pass": "green", "Fail": "red", "Blocked": "amber", "Not Run": "grey", "N/A": "dash" }[res] || "grey";
    return '<span class="poap-res poap-res-' + k + '"><em>' + label + "</em>" + esc(res) + "</span>";
  }
  function pathChips(nodes, active) {
    return nodes.map(function (n, i) {
      return (i ? '<span class="poap-arrow" aria-hidden="true">→</span>' : "") + '<button type="button" class="poap-sys' + (active && active.indexOf(n) >= 0 ? " on" : "") + '" data-poap-sys="' + esc(n) + '" title="Show every flow touching ' + esc(n) + '">' + esc(n) + "</button>";
    }).join("");
  }
  function flowDrawerHtml(f) {
    var r = '<div class="poap-dr-top"><span class="poap-dr-type"><i class="poap-sw" style="background:#06b6d4"></i>Integration flow</span><span class="poap-dr-id">' + esc(f.id) + "</span></div>";
    r += '<h3 class="poap-dr-title">' + esc(f.subProcess || f.topic || f.path) + "</h3>";
    r += '<div class="poap-dr-path">' + esc([f.topic, f.testNo].filter(Boolean).join(" › ")) + "</div>";
    r += '<div class="poap-dr-chips">' + chipHtml(f.status, statusKind(f.status)) + "</div>";
    r += '<div class="poap-dr-sec"><h4>System path</h4><div class="poap-pathrow">' + pathChips(f.pathNodes || [], null) + "</div></div>";
    r += '<div class="poap-dr-sec"><h4>Results</h4><div class="poap-resrow">' + resChip("2xx", f.r2xx) + resChip("3xx", f.r3xx) + resChip("4xx", f.r4xx) + resChip("5xx", f.r5xx) + "</div></div>";
    r += '<div class="poap-facts">' + fact("Planned date", f.dn != null ? fmtD(f.dn) : "TBC") + fact("Test no", esc(f.testNo)) + fact("Parent test", esc(f.parent)) + fact("Document", esc(f.doc)) + fact("Workbook row", f.row ? "P2P_Matrix!" + f.row : "") + "</div>";
    if (f.notes) r += '<div class="poap-dr-sec"><h4>Notes</h4><div class="poap-prose">' + paraHtml(f.notes.replace(/;\s+/g, " | ")) + "</div></div>";
    return r;
  }
  var groupStore = {};
  function groupDrawerHtml(g) {
    var r = '<div class="poap-dr-top"><span class="poap-dr-type"><i class="poap-sw" style="background:' + (g.color || "#64748b") + '"></i>' + esc(g.phase) + " · " + plural(g.bars.length, "activity").replace("activitys", "activities") + "</span></div>";
    r += '<h3 class="poap-dr-title">' + esc(g.topic) + "</h3>";
    r += '<div class="poap-dr-path">' + esc(g.pillar) + "</div>";
    r += '<div class="poap-dr-chips">' + chipHtml(g.status, statusKind(g.status)) + (g.rag ? chipHtml("RAG " + g.rag, g.rag.toLowerCase()) : "") + "</div>";
    r += '<div class="poap-dr-prog"><div class="poap-dr-progh"><span>Progress (duration-weighted)</span><b>' + Math.round(g.pct * 100) + '%</b></div><div class="poap-meter"><i style="width:' + Math.round(g.pct * 100) + '%;background:' + (g.color || "#64748b") + '"></i></div></div>';
    r += '<div class="poap-facts">' + fact("Start", g.s != null ? fmtD(g.s) : "TBC") + fact("End", g.s != null ? fmtD(g.e) : "TBC") + "</div>";
    r += '<div class="poap-dr-sec"><h4>Activities</h4><div class="poap-links">' + g.bars.map(function (b) { return linkBar(b.id); }).join("") + "</div></div>";
    return r;
  }
  function openDrawer(kind, id, group) {
    var M = S.model; if (!M) return;
    ensureDrawer();
    var html = "";
    if (kind === "bar") { var b = M.byId[id]; if (!b) return; html = barDrawerHtml(b); }
    else if (kind === "ms") { var m = M.msById[id]; if (!m) return; html = msDrawerHtml(m); }
    else if (kind === "flow") { var f = M.flowById[id]; if (!f) return; html = flowDrawerHtml(f); }
    else if (kind === "group") { var g = group || groupStore[id]; if (!g) return; html = groupDrawerHtml(g); }
    else return;
    if (!drawerEl.classList.contains("open")) drawerPrevFocus = document.activeElement;
    drawerBody.innerHTML = html;
    drawerBody.scrollTop = 0;
    drawerEl.classList.add("open");
    drawerEl.setAttribute("aria-label", "Details: " + (kind === "bar" ? M.byId[id].item : id));
    S.selKind = kind; S.selId = id;
    document.body.classList.add("poap-drawer-open");
    var x = drawerEl.querySelector(".poap-drawer-x");
    try { x.focus({ preventScroll: true }); } catch (e) { /* ignore */ }
    if (S.views.roadmap && S.view === "roadmap") S.views.roadmap.onSelect();
    markSelected();
  }
  function closeDrawer() {
    if (!drawerEl) return;
    drawerEl.classList.remove("open");
    document.body.classList.remove("poap-drawer-open");
    S.selKind = null; S.selId = null;
    markSelected();
    if (S.views.roadmap) S.views.roadmap.onSelect();
    if (drawerPrevFocus && document.contains(drawerPrevFocus)) { try { drawerPrevFocus.focus({ preventScroll: true }); } catch (e) { /* ignore */ } }
  }
  function markSelected() {
    if (!S.root) return;
    var sel = S.selKind + ":" + S.selId;
    Array.prototype.forEach.call(document.querySelectorAll(".poap-root [data-sel]"), function (el) {
      cls(el, "is-sel", el.getAttribute("data-sel") === sel);
    });
  }

  // ---------------------------------------------------------------
  // Shared UI helpers
  // ---------------------------------------------------------------
  var colCache = {};
  function colVars(hex) {
    if (colCache[hex]) return colCache[hex];
    var s = "--c:" + hex + ";--bg:" + hexToRgba(hex, 0.17) + ";--bd:" + hexToRgba(hex, 0.6) + ";--fl:" + hexToRgba(hex, 0.5) + ";--dk:" + darken(hex, 0.55) + ";";
    colCache[hex] = s; return s;
  }
  function tintVars(hex) {
    return "--t07:" + hexToRgba(hex, 0.07) + ";--t09:" + hexToRgba(hex, 0.09) + ";--t11:" + hexToRgba(hex, 0.11) + ";--t16:" + hexToRgba(hex, 0.16) + ";--t22:" + hexToRgba(hex, 0.22) + ";--t45:" + hexToRgba(hex, 0.45) + ";--t55:" + hexToRgba(hex, 0.55) + ";";
  }
  function segmented(items, active, onPick, label) {
    var wrap = h("div", { class: "toggle-pill poap-seg", role: "group", "aria-label": label || "Options" });
    items.forEach(function (it) {
      var b = h("button", { type: "button", "data-k": it.key, "aria-pressed": it.key === active ? "true" : "false", title: it.title || it.label, text: it.label });
      if (it.key === active) b.classList.add("active");
      b.addEventListener("click", function () { onPick(it.key); });
      wrap.appendChild(b);
    });
    wrap.setActive = function (k) {
      Array.prototype.forEach.call(wrap.children, function (b) {
        var on = b.getAttribute("data-k") === k; cls(b, "active", on); b.setAttribute("aria-pressed", on ? "true" : "false");
      });
    };
    return wrap;
  }
  var ICON = {
    chev: '<svg viewBox="0 0 16 16" width="12" height="12" aria-hidden="true"><path d="M5 3l5 5-5 5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    search: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/></svg>',
    fit: '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/></svg>',
    today: '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3"/></svg>',
    expand: '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M7 9l5-5 5 5M7 15l5 5 5-5"/></svg>',
    collapse: '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M7 4l5 5 5-5M7 20l5-5 5 5"/></svg>',
    print: '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V3h12v6M6 18H4v-7h16v7h-2M7 14h10v7H7z"/></svg>',
    pdf: '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 00-2 2v14a2 2 0 002 2h10a2 2 0 002-2V8z"/><path d="M14 3v5h5M12 11v6M9 14l3 3 3-3"/></svg>',
  };
  function emptyState(title, text) {
    return '<div class="poap-empty"><div class="poap-empty-ic"><svg viewBox="0 0 24 24" width="34" height="34" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="3"/><path d="M3 9h18M8 4v5M16 4v5M7 14h4M13 14h4"/></svg></div><h3>' + esc(title) + "</h3><p>" + esc(text) + "</p></div>";
  }

  // ===============================================================
  // ROADMAP (Gantt)
  // ===============================================================
  var RM = {
    labelW: 268, headH: 58, rowH: 28, pillarH: 38, collH: 32, msH: 20, pad: 10,
    layout: null, rendered: { items: {}, rows: {}, labels: {}, ghosts: {} }, geo: {}, ready: false, printing: false
  };
  var ZOOMS = { week: 6.5, month: 2.6, quarter: 0.95 };

  function rmMount(container) {
    var el = h("div", { class: "poap-rm" });
    // toolbar
    var tb = h("div", { class: "poap-rm-tb" });
    RM.zoomSeg = segmented([{ key: "week", label: "Week" }, { key: "month", label: "Month" }, { key: "quarter", label: "Quarter" }], S.zoom, function (k) { rmSetZoom(k); }, "Zoom level");
    RM.btnFit = h("button", { type: "button", class: "btn btn-outline btn-sm", title: "Fit the whole plan on screen", html: ICON.fit + "Fit to screen" });
    RM.btnFit.addEventListener("click", function () { rmFit(); });
    RM.btnToday = h("button", { type: "button", class: "btn btn-outline btn-sm", title: "Scroll to today", html: ICON.today + "Today" });
    RM.btnToday.addEventListener("click", function () { rmToday(true); });
    RM.btnExp = h("button", { type: "button", class: "btn btn-outline btn-sm", title: "Expand all lanes", html: ICON.expand + "Expand" });
    RM.btnExp.addEventListener("click", function () { rmSetAll(false); });
    RM.btnCol = h("button", { type: "button", class: "btn btn-outline btn-sm", title: "Collapse all topic lanes", html: ICON.collapse + "Collapse" });
    RM.btnCol.addEventListener("click", function () { rmSetAll(true); });
    RM.btnPrint = h("button", { type: "button", class: "btn btn-outline btn-sm poap-noprint", title: "Save the roadmap as a PDF", html: ICON.pdf + "Export PDF" });
    RM.btnPrint.addEventListener("click", function () { if (window.ATS && ATS.exportPdf) ATS.exportPdf(); });
    RM.count = h("div", { class: "poap-rm-count" });
    var grpA = h("div", { class: "poap-tb-grp" }, [RM.zoomSeg, RM.btnFit, RM.btnToday]);
    var grpB = h("div", { class: "poap-tb-grp" }, [RM.btnExp, RM.btnCol, RM.btnPrint]);
    tb.appendChild(grpA); tb.appendChild(grpB); tb.appendChild(RM.count);
    el.appendChild(tb);
    // minimap
    RM.mm = h("div", { class: "poap-mm", "aria-label": "Timeline minimap: drag the window to move along the plan", role: "group" });
    RM.mmCanvas = h("canvas", { class: "poap-mm-c" });
    RM.mmWin = h("div", { class: "poap-mm-win", tabindex: "0", role: "slider", "aria-label": "Visible window" });
    RM.mm.appendChild(RM.mmCanvas); RM.mm.appendChild(RM.mmWin);
    el.appendChild(RM.mm);
    // chart
    RM.card = h("div", { class: "poap-gcard card" });
    RM.sc = h("div", { class: "poap-gscroll", tabindex: "0", role: "region", "aria-label": "Roadmap timeline. Drag to pan, Ctrl+scroll to zoom." });
    RM.canvas = h("div", { class: "poap-gcanvas" });
    RM.head = h("div", { class: "poap-ghead" });
    RM.corner = h("div", { class: "poap-gcorner", html: "<span>Pillar › Topic</span>" });
    RM.hdTrack = h("div", { class: "poap-ghd-track" });
    RM.head.appendChild(RM.corner); RM.head.appendChild(RM.hdTrack);
    RM.body = h("div", { class: "poap-gbody" });
    RM.labels = h("div", { class: "poap-glabels" });
    RM.track = h("div", { class: "poap-gtrack" });
    RM.grid = h("div", { class: "poap-ggrid" });
    RM.rows = h("div", { class: "poap-grows" });
    RM.items = h("div", { class: "poap-gitems" });
    RM.svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    RM.svg.setAttribute("class", "poap-gdeps"); RM.svg.setAttribute("aria-hidden", "true");
    RM.svg.innerHTML = '<defs><marker id="poap-ah" viewBox="0 0 10 10" refX="8.5" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 1l8 4-8 4z" fill="#0f1c3f"/></marker>' +
      '<marker id="poap-ah2" viewBox="0 0 10 10" refX="8.5" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M0 1l8 4-8 4z" fill="#0ea5a0"/></marker></defs><g class="poap-deps-g"></g>';
    RM.depG = RM.svg.querySelector(".poap-deps-g");
    RM.lines = h("div", { class: "poap-glines" });
    RM.track.appendChild(RM.grid); RM.track.appendChild(RM.rows); RM.track.appendChild(RM.lines); RM.track.appendChild(RM.svg); RM.track.appendChild(RM.items);
    RM.hl = h("div", { class: "poap-hl", "aria-hidden": "true" });
    RM.body.appendChild(RM.labels); RM.body.appendChild(RM.track); RM.body.appendChild(RM.hl);
    RM.canvas.appendChild(RM.head); RM.canvas.appendChild(RM.body);
    RM.sc.appendChild(RM.canvas);
    RM.card.appendChild(RM.sc);
    RM.empty = h("div", { class: "poap-gempty", style: "display:none" });
    RM.card.appendChild(RM.empty);
    el.appendChild(RM.card);
    // legend
    RM.legend = h("div", { class: "poap-legend card" });
    el.appendChild(RM.legend);
    container.appendChild(el);
    RM.el = el;
    rmBind();
    RM.ready = true;
  }

  function rmBind() {
    var sc = RM.sc;
    sc.addEventListener("scroll", rafThrottle(function () { rmDraw(); rmUpdateWin(); hideTip(); }), { passive: true });
    // click / hover (delegated)
    RM.items.addEventListener("click", function (ev) {
      if (RM.dragMoved) return;
      var b = ev.target.closest(".poap-bar"); if (b) { openDrawer("bar", b.getAttribute("data-id")); return; }
      var m = ev.target.closest(".poap-ms"); if (m) openDrawer("ms", m.getAttribute("data-id"));
    });
    RM.items.addEventListener("keydown", function (ev) {
      if (ev.key !== "Enter" && ev.key !== " ") return;
      var b = ev.target.closest(".poap-bar"); if (b) { ev.preventDefault(); openDrawer("bar", b.getAttribute("data-id")); return; }
      var m = ev.target.closest(".poap-ms"); if (m) { ev.preventDefault(); openDrawer("ms", m.getAttribute("data-id")); }
    });
    RM.items.addEventListener("mouseover", function (ev) {
      var b = ev.target.closest(".poap-bar");
      if (b) {
        var id = b.getAttribute("data-id");
        if (S.hoverId !== id) { S.hoverId = id; rmDrawDeps(); }
        showTip(barTipHtml(S.model.byId[id]), ev); return;
      }
      var m = ev.target.closest(".poap-ms");
      if (m) { showTip(msTipHtml(S.model.msById[m.getAttribute("data-id")]), ev); return; }
      var s = ev.target.closest(".poap-bsum");
      if (s) { showTip(sumTipHtml(s), ev); }
    });
    RM.items.addEventListener("mousemove", function (ev) { moveTip(ev); });
    RM.items.addEventListener("mouseout", function (ev) {
      var b = ev.target.closest(".poap-bar");
      if (b && S.hoverId) { S.hoverId = null; rmDrawDeps(); }
      if (!ev.relatedTarget || !ev.relatedTarget.closest || !ev.relatedTarget.closest(".poap-bar,.poap-ms,.poap-bsum")) hideTip();
    });
    RM.items.addEventListener("focusin", function (ev) {
      var b = ev.target.closest(".poap-bar"); if (b) { S.hoverId = b.getAttribute("data-id"); rmDrawDeps(); }
    });
    RM.items.addEventListener("focusout", function () { if (S.hoverId) { S.hoverId = null; rmDrawDeps(); } });
    // row highlight: hover anywhere on an activity / item row (label or timeline) to light up the whole row
    RM.body.addEventListener("mousemove", rafThrottle(function (ev) { rmHover(ev); }));
    RM.body.addEventListener("mouseleave", function () { rmHl(null); });
    RM.labels.addEventListener("mouseover", function (ev) {
      var it = ev.target.closest(".poap-rl-item");
      if (it) { var id = it.getAttribute("data-id"); if (S.hoverId !== id) { S.hoverId = id; rmDrawDeps(); } showTip(barTipHtml(S.model.byId[id]), ev); return; }
      var ac = ev.target.closest(".poap-rl-act");
      if (ac) { var tp = ac.querySelector(".poap-rl-name"); showTip('<div class="poap-tip-t">' + esc(tp ? tp.textContent : "") + " — " + esc(ac.querySelector(".poap-rl-meta").textContent) + "</div>", ev); }
    });
    RM.labels.addEventListener("mousemove", function (ev) { moveTip(ev); });
    RM.labels.addEventListener("mouseout", function (ev) {
      if (S.hoverId && ev.target.closest(".poap-rl-item")) { S.hoverId = null; rmDrawDeps(); }
      if (!ev.relatedTarget || !ev.relatedTarget.closest || !ev.relatedTarget.closest(".poap-rl")) hideTip();
    });
    RM.labels.addEventListener("click", function (ev) {
      var it = ev.target.closest(".poap-rl-item"); if (it) openDrawer("bar", it.getAttribute("data-id"));
    });
    // labels: collapse/expand
    RM.labels.addEventListener("click", function (ev) {
      var l = ev.target.closest(".poap-lab"); if (!l) return;
      rmToggle(l.getAttribute("data-key"));
    });
    RM.labels.addEventListener("keydown", function (ev) {
      if (ev.key !== "Enter" && ev.key !== " ") return;
      var l = ev.target.closest(".poap-lab"); if (!l) return; ev.preventDefault(); rmToggle(l.getAttribute("data-key"));
    });
    // drag to pan
    var drag = null;
    sc.addEventListener("pointerdown", function (ev) {
      if (ev.button !== 0) return;
      if (ev.target.closest(".poap-bar,.poap-ms,.poap-lab,.poap-rl,button,a,.poap-bsum")) return;
      drag = { x: ev.clientX, y: ev.clientY, sl: sc.scrollLeft, st: sc.scrollTop, id: ev.pointerId };
      RM.dragMoved = false;
    });
    sc.addEventListener("pointermove", function (ev) {
      if (!drag) return;
      var dx = ev.clientX - drag.x, dy = ev.clientY - drag.y;
      if (!RM.dragMoved && Math.abs(dx) + Math.abs(dy) < 4) return;
      if (!RM.dragMoved) { RM.dragMoved = true; sc.classList.add("is-drag"); try { sc.setPointerCapture(drag.id); } catch (e) { /* ignore */ } }
      sc.scrollLeft = drag.sl - dx; sc.scrollTop = drag.st - dy;
    });
    function endDrag(ev) {
      if (!drag) return;
      try { sc.releasePointerCapture(drag.id); } catch (e) { /* ignore */ }
      drag = null; sc.classList.remove("is-drag");
      setTimeout(function () { RM.dragMoved = false; }, 0);
    }
    sc.addEventListener("pointerup", endDrag); sc.addEventListener("pointercancel", endDrag);
    // ctrl/cmd + wheel -> zoom around pointer
    sc.addEventListener("wheel", function (ev) {
      if (!(ev.ctrlKey || ev.metaKey)) return;
      ev.preventDefault();
      var f = ev.deltaY < 0 ? 1.18 : 1 / 1.18;
      rmZoomTo(clamp(S.pxd * f, 0.15, 14), ev.clientX);
    }, { passive: false });
    // minimap
    var mmDrag = null;
    RM.mm.addEventListener("pointerdown", function (ev) {
      if (ev.button !== 0) return;
      var r = RM.mm.getBoundingClientRect();
      var onWin = ev.target === RM.mmWin;
      if (!onWin) { rmMmCenter((ev.clientX - r.left) / r.width); }
      mmDrag = { id: ev.pointerId, off: onWin ? (ev.clientX - RM.mmWin.getBoundingClientRect().left) : RM.mmWin.offsetWidth / 2 };
      try { RM.mm.setPointerCapture(ev.pointerId); } catch (e) { /* ignore */ }
      RM.mm.classList.add("is-drag");
      ev.preventDefault();
    });
    RM.mm.addEventListener("pointermove", function (ev) {
      if (!mmDrag) return;
      var r = RM.mm.getBoundingClientRect();
      var left = (ev.clientX - r.left - mmDrag.off) / r.width;
      rmMmLeft(left);
    });
    function mmEnd() { if (!mmDrag) return; try { RM.mm.releasePointerCapture(mmDrag.id); } catch (e) { /* ignore */ } mmDrag = null; RM.mm.classList.remove("is-drag"); }
    RM.mm.addEventListener("pointerup", mmEnd); RM.mm.addEventListener("pointercancel", mmEnd);
    RM.mmWin.addEventListener("keydown", function (ev) {
      var step = (ev.shiftKey ? 400 : 120);
      if (ev.key === "ArrowRight") { RM.sc.scrollLeft += step; ev.preventDefault(); }
      else if (ev.key === "ArrowLeft") { RM.sc.scrollLeft -= step; ev.preventDefault(); }
    });
    // print: render everything
    window.addEventListener("beforeprint", function () { if (!RM.ready || !S.model) return; RM.printing = true; rmDraw(true); });
    window.addEventListener("afterprint", function () { RM.printing = false; if (RM.ready && S.model) { clearRendered(); rmDraw(); } });
  }

  // ---- row hover highlight -----------------------------------------
  function rmHl(box) {
    if (!RM.hl) return;
    if (!box) { RM.hl.style.display = "none"; RM.hlKey = null; return; }
    var k = box.top + ":" + box.h; if (RM.hlKey === k) return; RM.hlKey = k;
    RM.hl.style.cssText = "display:block;top:" + box.top + "px;height:" + box.h + "px;--hc:" + (box.color || "#0ea5a0") + ";" + tintVars(box.color || "#0ea5a0");
    RM.hl.className = "poap-hl" + (box.group ? " is-group" : "");
  }
  function rmHover(ev) {
    if (!RM.layout || RM.dragMoved) return;
    var y = ev.clientY - RM.body.getBoundingClientRect().top, lanes = RM.layout.lanes, ln = null, i;
    for (i = 0; i < lanes.length; i++) { if (y >= lanes[i].y && y < lanes[i].y + lanes[i].h) { ln = lanes[i]; break; } }
    if (!ln || ln.kind !== "topic" || ln.collapsed || !ln.rowsList || y < ln.y + ln.msTop) { rmHl(null); return; }
    var idx = Math.floor((y - ln.y - ln.msTop) / RM.rowH), rw = ln.rowsList[idx];
    if (!rw) { rmHl(null); return; }
    var top = ln.y + ln.msTop + idx * RM.rowH;
    if (rw.kind === "act") rmHl({ top: top, h: (rw.n + 1) * RM.rowH, color: rw.color, group: true });
    else rmHl({ top: top, h: RM.rowH, color: rw.b.color });
  }

  // ---- layout -----------------------------------------------------
  function weightedPct(list) {
    var num = 0, den = 0;
    list.forEach(function (b) { var w = b.dur || 14; num += b.pct * w; den += w; });
    return den ? num / den : 0;
  }
  function worstRag(list) {
    var r = "";
    list.forEach(function (b) { if (b.rag === "Red") r = "Red"; else if (b.rag === "Amber" && r !== "Red") r = "Amber"; else if (b.rag === "Green" && !r) r = "Green"; });
    return r;
  }
  function packLane(bars) {
    var sorted = bars.slice().sort(function (a, b) { return a.ps - b.ps || a.pe - b.pe; });
    var ends = [];
    sorted.forEach(function (b) {
      var r = 0;
      while (r < ends.length && ends[r] >= b.ps - 1) r++;
      ends[r] = b.pe; b.sub = r;
    });
    return Math.max(1, ends.length);
  }

  function msLabel(m, ln) {
    var lab = m.name;
    if (ln.kind === "topic" && m.topic && lab.length > m.topic.length && lab.slice(-m.topic.length) === m.topic) lab = lab.slice(0, -m.topic.length).replace(/\s*[-–·]\s*$/, "");
    return lab;
  }
  function rmComputeLayout() {
    var M = S.model, pxd = S.pxd, minDn = M.minDn;
    var bars = S.fb, ms = S.fm;
    var pillarOrder = M.pillars.map(function (p) { return p.name; });
    var byPillar = {}, order = [];
    bars.forEach(function (b) {
      if (!byPillar[b.pillar]) { byPillar[b.pillar] = { name: b.pillar, topics: {}, topicOrder: [], bars: [], ms: [] }; }
      var P = byPillar[b.pillar];
      if (!P.topics[b.topic]) { P.topics[b.topic] = { name: b.topic, bars: [], ms: [], first: b.idx }; P.topicOrder.push(b.topic); }
      P.topics[b.topic].bars.push(b); P.bars.push(b);
    });
    ms.forEach(function (m) {
      var pn = m.pillar || "Cross-cutting";
      if (!byPillar[pn]) byPillar[pn] = { name: pn, topics: {}, topicOrder: [], bars: [], ms: [] };
      var P = byPillar[pn];
      if (m.topic && P.topics[m.topic]) P.topics[m.topic].ms.push(m); else P.ms.push(m);
    });
    // pillar-level milestone filter: skip when a bar filter is active and pillar has no bars
    var names = Object.keys(byPillar).sort(function (a, b) {
      var ia = M.pillarIdx[a], ib = M.pillarIdx[b];
      return (ia == null ? 99 : ia) - (ib == null ? 99 : ib);
    });
    var Y = 0, lanes = [], rows = [];
    names.forEach(function (pn) {
      var P = byPillar[pn];
      if (barFiltersActive() && !P.bars.length) return;
      var pcol = M.pillarColor[pn] || "#64748b";
      var pkey = "p:" + pn, pcoll = !!S.collapsed[pkey];
      var prow = { kind: "pillar", key: pkey, name: pn, color: pcol, y: Y, h: RM.pillarH, bars: P.bars, ms: P.ms.slice(), collapsed: pcoll, n: P.bars.length, pct: weightedPct(P.bars), rag: worstRag(P.bars) };
      if (P.bars.length) {
        prow.smin = Math.min.apply(null, P.bars.map(function (b) { return b.ps; }));
        prow.smax = Math.max.apply(null, P.bars.map(function (b) { return b.pe; }));
      }
      lanes.push(prow); Y += prow.h;
      if (pcoll) {
        // milestones of collapsed topics roll up to the pillar row
        P.topicOrder.forEach(function (tn) { prow.ms = prow.ms.concat(P.topics[tn].ms); });
        return;
      }
      var tnames = P.topicOrder.slice().sort(function (a, b) { return P.topics[a].first - P.topics[b].first; });
      tnames.forEach(function (tn) {
        var T = P.topics[tn], tkey = "t:" + pn + "|" + tn, tcoll = !!S.collapsed[tkey];
        var lane = { kind: "topic", key: tkey, name: tn, pillar: pn, color: pcol, y: Y, bars: T.bars, ms: T.ms, collapsed: tcoll, n: T.bars.length, pct: weightedPct(T.bars), rag: worstRag(T.bars), items: [] };
        if (tcoll) { lane.h = RM.collH; lane.nsub = 1; lane.msTop = 0; }
        else {
          // expanded: a summary strip (topic span + milestones), then one row per activity type and one row per item under it
          var groups = {}, gorder = [], rowsList = [], rr = 0;
          T.bars.forEach(function (b) {
            if (!groups[b.type]) { groups[b.type] = { type: b.type, color: b.color, bars: [] }; gorder.push(b.type); }
            groups[b.type].bars.push(b);
          });
          gorder.sort(function (a, b) { return Math.min.apply(null, groups[a].bars.map(function (x) { return x.ps; })) - Math.min.apply(null, groups[b].bars.map(function (x) { return x.ps; })); });
          gorder.forEach(function (tn2) {
            var g = groups[tn2];
            g.bars.sort(function (a, b) { return a.ps - b.ps || a.pe - b.pe; });
            rowsList.push({ kind: "act", type: g.type, color: g.color, n: g.bars.length, pct: weightedPct(g.bars), rag: worstRag(g.bars), smin: Math.min.apply(null, g.bars.map(function (x) { return x.ps; })), smax: Math.max.apply(null, g.bars.map(function (x) { return x.pe; })), row: rr });
            rr++;
            g.bars.forEach(function (b) { b.sub = rr; rowsList.push({ kind: "item", b: b, row: rr }); rr++; });
          });
          lane.rowsList = rowsList; lane.nsub = rr;
          lane.msTop = RM.collH;
          lane.h = lane.msTop + rr * RM.rowH + 8;
        }
        T.bars.forEach(function (b) {
          lane.items.push({ b: b, x: (b.ps - minDn) * pxd, w: Math.max((b.pe - b.ps + 1) * pxd, 5), row: tcoll ? 0 : b.sub });
        });
        lane.items.sort(function (a, b) { return a.x - b.x; });
        lanes.push(lane); Y += lane.h;
      });
    });
    // geometry map (for dependency arrows and focus)
    var geo = {};
    lanes.forEach(function (ln) {
      if (ln.kind === "topic") {
        ln.items.forEach(function (it) {
          var cy = ln.collapsed ? ln.y + ln.h / 2 : ln.y + ln.msTop + it.row * RM.rowH + 4 + 10;
          geo[it.b.id] = { x: it.x, w: it.w, cy: cy, lane: ln };
        });
      }
    });
    // bars hidden in collapsed pillars -> pillar row centre
    lanes.forEach(function (ln) {
      if (ln.kind === "pillar" && ln.collapsed) ln.bars.forEach(function (b) { geo[b.id] = { x: (b.ps - minDn) * pxd, w: Math.max((b.pe - b.ps + 1) * pxd, 5), cy: ln.y + ln.h / 2, lane: ln }; });
    });
    // milestone label visibility (greedy, per lane)
    lanes.forEach(function (ln) {
      var list = ln.ms.slice().sort(function (a, b) { return a.dn - b.dn; });
      var lastEnd = -1e9;
      list.forEach(function (m) {
        m._x = (m.dn - minDn) * pxd;
        m._lab = msLabel(m, ln);
        var wLab = Math.min(190, m._lab.length * 5.6 + 8);
        m._label = (m._x - 7 >= lastEnd);
        if (m._label) lastEnd = m._x + 12 + wLab; else lastEnd = Math.max(lastEnd, m._x + 8);
        m._lw = wLab;
      });
      ln.ms = list;
    });
    RM.layout = { lanes: lanes, totalH: Y + 16, geo: geo, tlW: (M.maxDn - M.minDn + 1) * pxd };
    RM.geo = geo;
  }

  function rmBuildHeader() {
    var M = S.model, pxd = S.pxd, minDn = M.minDn, W = RM.layout.tlW;
    var mode = pxd >= 3.5 ? "week" : (pxd >= 1.3 ? "month" : "quarter");
    RM.mode = mode;
    var a = [], b = [], g = [], bands = [];
    var d = dnParts(minDn), y = d.y, m = d.m;
    var lastYear = -1;
    // iterate months
    var cursor = dnFromYMD(y, m + 1, 1);
    var idx = 0;
    while (cursor <= M.maxDn) {
      var p = dnParts(cursor);
      var next = dnFromYMD(p.m === 11 ? p.y + 1 : p.y, p.m === 11 ? 1 : p.m + 2, 1);
      var x0 = (cursor - minDn) * pxd, x1 = (next - minDn) * pxd;
      var isQ = (p.m % 3 === 0);
      if (mode === "week") {
        a.push('<div class="poap-hcell poap-hA" style="left:' + x0 + "px;width:" + (x1 - x0) + 'px"><span>' + MON_FULL[p.m] + " " + p.y + "</span></div>");
        g.push('<i class="poap-gl poap-gl-m" style="left:' + x0 + 'px"></i>');
      } else if (mode === "month") {
        if (p.y !== lastYear) {
          var ny = dnFromYMD(p.y + 1, 1, 1);
          a.push('<div class="poap-hcell poap-hA" style="left:' + x0 + "px;width:" + (Math.min(ny, M.maxDn + 1) - cursor) * pxd + 'px"><span>' + p.y + "</span></div>"); lastYear = p.y;
        }
        b.push('<div class="poap-hcell poap-hB" style="left:' + x0 + "px;width:" + (x1 - x0) + 'px"><span>' + MON[p.m] + "</span></div>");
        g.push('<i class="poap-gl ' + (isQ ? "poap-gl-q" : "poap-gl-m") + '" style="left:' + x0 + 'px"></i>');
      } else {
        if (p.y !== lastYear) {
          var ny2 = dnFromYMD(p.y + 1, 1, 1);
          a.push('<div class="poap-hcell poap-hA" style="left:' + x0 + "px;width:" + (Math.min(ny2, M.maxDn + 1) - cursor) * pxd + 'px"><span>' + p.y + "</span></div>"); lastYear = p.y;
        }
        if (isQ) {
          var qn = dnFromYMD(p.m >= 9 ? p.y + 1 : p.y, p.m >= 9 ? 1 : p.m + 4, 1);
          b.push('<div class="poap-hcell poap-hB" style="left:' + x0 + "px;width:" + (Math.min(qn, M.maxDn + 1) - cursor) * pxd + 'px"><span>Q' + (p.m / 3 + 1) + "</span></div>");
          g.push('<i class="poap-gl poap-gl-q" style="left:' + x0 + 'px"></i>');
        }
      }
      if ((idx % 2) === 0) bands.push('<i class="poap-band" style="left:' + x0 + "px;width:" + (x1 - x0) + 'px"></i>');
      cursor = next; idx++;
    }
    if (mode === "week") {
      var mon = mondayOf(minDn); if (mon < minDn) mon += 7;
      var every = (pxd * 7 < 26) ? 2 : 1, k = 0;
      for (; mon <= M.maxDn; mon += 7, k++) {
        var x = (mon - minDn) * pxd, pp = dnParts(mon);
        if (k % every === 0) b.push('<div class="poap-hcell poap-hB poap-hW" style="left:' + x + "px;width:" + (7 * pxd * every) + 'px"><span>' + pad2(pp.d) + "</span></div>");
        g.push('<i class="poap-gl poap-gl-w" style="left:' + x + 'px"></i>');
      }
    }
    // holiday / freeze bands
    var hol = [];
    M.bars.forEach(function (bb) {
      if (!bb.isHol || bb.s == null) return;
      if (S.filters.pillar && bb.pillar !== S.filters.pillar && bb.pillar !== "Cross-cutting") return;
      var x = (bb.s - minDn) * pxd, w = Math.max((bb.e - bb.s + 1) * pxd, 3);
      hol.push('<div class="poap-hol" style="left:' + x + "px;width:" + w + 'px" title="' + esc(bb.item + " · " + fmtD(bb.s, true) + " – " + fmtD(bb.e)) + '"><span>' + esc(bb.item) + "</span></div>");
    });
    RM.hdTrack.style.width = W + "px";
    RM.hdTrack.innerHTML = a.join("") + b.join("");
    RM.grid.innerHTML = bands.join("") + g.join("") + hol.join("");
    // today / status lines
    var tx = (M.todayDn - minDn) * pxd, sx = (M.statusDn - minDn) * pxd;
    var lines = "", flags = "";
    var same = (M.todayDn === M.statusDn);
    lines += '<i class="poap-line poap-line-today" style="left:' + tx + 'px"></i>';
    flags += '<div class="poap-flag poap-flag-today" style="left:' + tx + 'px" title="' + (same ? "Today / status date " : "Today ") + fmtD(M.todayDn) + '"><span>' + (same ? "Today · " : "Today ") + fmtShort(M.todayDn) + "</span></div>";
    if (!same) {
      lines += '<i class="poap-line poap-line-status" style="left:' + sx + 'px"></i>';
      flags += '<div class="poap-flag poap-flag-status" style="left:' + sx + 'px" title="Status date ' + fmtD(M.statusDn) + '"><span>Status ' + fmtShort(M.statusDn) + "</span></div>";
    }
    RM.lines.innerHTML = lines;
    RM.hdTrack.insertAdjacentHTML("beforeend", flags);
  }

  function clearRendered() {
    RM.items.innerHTML = ""; RM.rows.innerHTML = ""; RM.labels.innerHTML = "";
    RM.rendered = { items: {}, rows: {}, labels: {}, ghosts: {} };
  }

  function rmLayoutAll(keepScroll) {
    if (!S.model || !RM.ready) return;
    var L = null;
    rmComputeLayout();
    L = RM.layout;
    var empty = !L.lanes.length;
    RM.empty.style.display = empty ? "" : "none";
    RM.sc.style.visibility = empty ? "hidden" : "";
    if (empty) {
      RM.empty.innerHTML = emptyState(S.model.bars.length ? "No activities match the filters" : "No activities in this workbook", S.model.bars.length ? "Try clearing a filter or the search box." : "Add rows to the POAP_Plan sheet and reload the workbook.");
    }
    RM.canvas.style.width = (RM.labelW + L.tlW) + "px";
    RM.body.style.height = L.totalH + "px";
    RM.body.style.width = (RM.labelW + L.tlW) + "px";
    RM.labels.style.height = L.totalH + "px";
    RM.track.style.width = L.tlW + "px"; RM.track.style.height = L.totalH + "px";
    RM.track.style.left = RM.labelW + "px";
    RM.svg.setAttribute("width", L.tlW); RM.svg.setAttribute("height", L.totalH);
    rmBuildHeader();
    clearRendered();
    rmDraw();
    rmDrawDeps();
    rmDrawMinimap();
    rmUpdateWin();
    RM.count.innerHTML = "<b>" + S.fb.length + "</b> activities · <b>" + L.lanes.filter(function (l) { return l.kind === "topic"; }).length + "</b> topic lanes · " + S.fm.length + " milestones";
    RM.zoomSeg.setActive(S.zoom);
    cls(RM.btnFit, "is-on", S.zoom === "fit");
  }

  // ---- drawing (virtualised) -------------------------------------
  function sumTipHtml(el) { return '<div class="poap-tip-t">' + esc(el.getAttribute("data-tip") || "") + "</div>"; }

  function makeBarEl(it, ln) {
    var b = it.b, M = S.model;
    var el = document.createElement("div");
    var c = ["poap-bar"];
    if (b.status === "Complete") c.push("is-done");
    if (b.tbc) c.push("is-tbc");
    if (b.slip > 0 || b.status === "Delayed") c.push("is-late");
    if (b.status === "On Hold") c.push("is-hold");
    if (it.w < 52) c.push("is-narrow");
    el.className = c.join(" ");
    el.setAttribute("data-id", b.id); el.setAttribute("data-sel", "bar:" + b.id);
    el.setAttribute("tabindex", "0"); el.setAttribute("role", "button");
    el.setAttribute("aria-label", b.type + ": " + b.item + ", " + (b.s != null ? fmtD(b.s) + " to " + fmtD(b.e) : "dates TBC") + ", " + b.status + ", " + Math.round(b.pct * 100) + " percent");
    var top = ln.y + ln.msTop + it.row * RM.rowH + 4;
    el.style.cssText = colVars(b.color) + "left:" + it.x + "px;top:" + top + "px;width:" + it.w + "px;";
    var html = '<i class="poap-bar-fill" style="width:' + Math.round(b.pct * 100) + '%"></i>';
    html += '<span class="poap-bar-lbl">' + (b.rag ? '<i class="poap-rag poap-rag-' + b.rag.toLowerCase() + '"></i>' : "") + (b.status === "Complete" ? '<b class="poap-tick">✓</b>' : "") + esc(b.item) + "</span>";
    if (b.slip > 0) html += '<span class="poap-bar-slip' + (b.slip > 14 ? " is-red" : "") + '">+' + b.slip + "d</span>";
    else if (b.slip < 0) html += '<span class="poap-bar-slip is-ahead">' + b.slip + "d</span>";
    el.innerHTML = html;
    if (S.selKind === "bar" && S.selId === b.id) el.classList.add("is-sel");
    return el;
  }
  function makeGhostEl(it, ln) {
    var b = it.b, M = S.model;
    var gx = (b.bs - M.minDn) * S.pxd, gw = Math.max((b.be - b.bs + 1) * S.pxd, 4);
    var top = ln.y + ln.msTop + it.row * RM.rowH + 4 + 21;
    var el = document.createElement("div");
    el.className = "poap-ghost"; el.style.cssText = "left:" + gx + "px;top:" + top + "px;width:" + gw + "px;";
    el.setAttribute("title", "Baseline " + fmtD(b.bs, true) + " – " + fmtD(b.be));
    return el;
  }
  function makeMsEl(m, ln, cy) {
    var el = document.createElement("div");
    el.className = "poap-ms" + (m._label ? " has-lbl" : "") + " poap-ms-" + m.status.toLowerCase().replace(/\s+/g, "");
    el.setAttribute("data-id", m.id); el.setAttribute("data-sel", "ms:" + m.id);
    el.setAttribute("tabindex", "0"); el.setAttribute("role", "button");
    el.setAttribute("aria-label", "Milestone: " + m.name + ", " + fmtD(m.dn));
    el.style.cssText = "left:" + (m._x - 7) + "px;top:" + (cy - 7) + "px;--mc:" + m.color + ";";
    el.innerHTML = '<i class="poap-dia"></i>' + (m._label ? '<span class="poap-ms-lbl" style="max-width:' + Math.round(m._lw) + 'px">' + esc(m._lab || m.name) + "</span>" : "");
    if (S.selKind === "ms" && S.selId === m.id) el.classList.add("is-sel");
    return el;
  }

  function lanePctRing(p, col) {
    var r = 7, C = 2 * Math.PI * r, o = C * (1 - p);
    return '<svg class="poap-ring" viewBox="0 0 18 18" width="18" height="18" aria-hidden="true"><circle cx="9" cy="9" r="' + r + '" fill="none" stroke="rgba(15,28,63,.1)" stroke-width="2.6"/><circle cx="9" cy="9" r="' + r + '" fill="none" stroke="' + col + '" stroke-width="2.6" stroke-linecap="round" stroke-dasharray="' + C.toFixed(2) + '" stroke-dashoffset="' + o.toFixed(2) + '" transform="rotate(-90 9 9)"/></svg>';
  }
  function makeLabelEl(ln) {
    var el = document.createElement("div");
    var open = !ln.collapsed;
    el.className = "poap-lab " + (ln.kind === "pillar" ? "poap-lab-p" : "poap-lab-t") + (open ? "" : " is-coll");
    el.setAttribute("data-key", ln.key); el.setAttribute("role", "button"); el.setAttribute("tabindex", "0");
    el.setAttribute("aria-expanded", open ? "true" : "false");
    el.setAttribute("aria-label", (open ? "Collapse " : "Expand ") + ln.name);
    el.style.cssText = "top:" + ln.y + "px;height:" + (ln.kind === "topic" && !ln.collapsed ? ln.msTop : ln.h) + "px;--pc:" + ln.color + ";";
    var rag = ln.rag ? '<i class="poap-rag poap-rag-' + ln.rag.toLowerCase() + '" title="Worst RAG: ' + ln.rag + '"></i>' : "";
    if (ln.kind === "pillar") {
      el.innerHTML = '<span class="poap-chev">' + ICON.chev + '</span><i class="poap-pdot"></i><span class="poap-lab-name">' + esc(ln.name) + '</span><span class="poap-lab-meta">' + ln.n + " · " + Math.round(ln.pct * 100) + "%</span>" + lanePctRing(ln.pct, ln.color);
    } else {
      el.innerHTML = '<span class="poap-chev">' + ICON.chev + '</span><div class="poap-lab-main"><div class="poap-lab-name" title="' + esc(ln.name) + '">' + esc(ln.name) + '</div><div class="poap-lab-sub">' + plural(ln.n, "activity").replace("activitys", "activities") + " · " + Math.round(ln.pct * 100) + "% " + rag + "</div></div>";
    }
    return el;
  }

  function rmDraw(all) {
    if (!S.model || !RM.layout) return;
    var L = RM.layout, sc = RM.sc;
    var vh = sc.clientHeight || 600, vw = sc.clientWidth || 1000;
    var sy = sc.scrollTop, sx = sc.scrollLeft;
    var y0 = sy - RM.headH - 160, y1 = sy + vh + 160;
    var x0 = sx - RM.labelW - 400, x1 = sx + vw + 400;
    if (all || RM.printing) { y0 = -1e9; y1 = 1e9; x0 = -1e9; x1 = 1e9; }
    var used = { items: {}, rows: {}, labels: {}, ghosts: {} };
    var R = RM.rendered, minDn = S.model.minDn, pxd = S.pxd;
    var fragI = document.createDocumentFragment(), fragR = document.createDocumentFragment(), fragL = document.createDocumentFragment();
    var zebra = 0;
    L.lanes.forEach(function (ln) {
      if (ln.kind === "topic") zebra++;
      if (ln.y + ln.h < y0 || ln.y > y1) return;
      // row stripe
      var rk = ln.key;
      used.rows[rk] = 1;
      if (!R.rows[rk]) {
        var rs = document.createElement("div");
        rs.className = "poap-row " + (ln.kind === "pillar" ? "poap-row-p" : "poap-row-t" + (zebra % 2 ? " odd" : ""));
        rs.style.cssText = "top:" + ln.y + "px;height:" + ln.h + "px;--pc:" + ln.color + ";";
        R.rows[rk] = rs; fragR.appendChild(rs);
      }
      if (!R.labels[rk]) { var lb = makeLabelEl(ln); R.labels[rk] = lb; fragL.appendChild(lb); }
      used.labels[rk] = 1;
      // pillar summary
      if (ln.kind === "pillar" && ln.n) {
        var sk = "ps:" + ln.key; used.items[sk] = 1;
        if (!R.items[sk]) {
          var sx0 = (ln.smin - minDn) * pxd, sw = Math.max((ln.smax - ln.smin + 1) * pxd, 6);
          {
            var se = document.createElement("div");
            se.className = "poap-bsum poap-psum";
            se.setAttribute("data-tip", ln.name + " · " + plural(ln.n, "activity").replace("activitys", "activities") + " · " + Math.round(ln.pct * 100) + "% complete · " + fmtD(ln.smin, true) + " – " + fmtD(ln.smax));
            se.style.cssText = "left:" + sx0 + "px;top:" + (ln.y + 12) + "px;width:" + sw + "px;--pc:" + ln.color + ";";
            se.innerHTML = '<i style="width:' + Math.round(ln.pct * 100) + '%"></i>';
            R.items[sk] = se; fragI.appendChild(se);
          }
        }
      }
      // collapsed topic summary
      if (ln.kind === "topic" && ln.collapsed) {
        ln.items.forEach(function (it) {
          if (it.x + it.w < x0 || it.x > x1) return;
          var ck = "cs:" + it.b.id; used.items[ck] = 1;
          if (!R.items[ck]) {
            var ce = document.createElement("div");
            ce.className = "poap-bsum" + (it.b.tbc ? " is-tbc" : "");
            ce.setAttribute("data-tip", it.b.type + ": " + it.b.item + (it.b.s != null ? " · " + fmtD(it.b.s, true) + " – " + fmtD(it.b.e) : " · dates TBC"));
            ce.style.cssText = colVars(it.b.color) + "left:" + it.x + "px;top:" + (ln.y + 11) + "px;width:" + it.w + "px;";
            R.items[ck] = ce; fragI.appendChild(ce);
          }
        });
      } else if (ln.kind === "topic") {
        // topic span strip
        if (ln.bars.length) {
          var tk = "ts:" + ln.key; used.items[tk] = 1;
          if (!R.items[tk]) {
            var tmin = Math.min.apply(null, ln.bars.map(function (b) { return b.ps; })), tmax = Math.max.apply(null, ln.bars.map(function (b) { return b.pe; }));
            var te = document.createElement("div"); te.className = "poap-bsum poap-tsum";
            te.setAttribute("data-tip", ln.name + " · " + plural(ln.n, "activity").replace("activitys", "activities") + " · " + Math.round(ln.pct * 100) + "% complete · " + fmtD(tmin, true) + " – " + fmtD(tmax));
            te.style.cssText = "left:" + ((tmin - minDn) * pxd) + "px;top:" + (ln.y + ln.msTop - 9) + "px;width:" + Math.max((tmax - tmin + 1) * pxd, 6) + "px;--pc:" + ln.color + ";";
            te.innerHTML = '<i style="width:' + Math.round(ln.pct * 100) + '%"></i>';
            R.items[tk] = te; fragI.appendChild(te);
          }
        }
        // one label + guide row per activity header / item
        (ln.rowsList || []).forEach(function (rw) {
          var ry = ln.y + ln.msTop + rw.row * RM.rowH;
          if (ry + RM.rowH < y0 || ry > y1) return;
          var lk = "rl:" + ln.key + ":" + rw.row; used.labels[lk] = 1;
          if (!R.labels[lk]) {
            var le = document.createElement("div");
            le.className = "poap-rl poap-rl-" + rw.kind; le.style.cssText = "top:" + ry + "px;height:" + RM.rowH + "px;" + (rw.kind === "act" ? "--tc:" + rw.color + ";" + tintVars(rw.color) : colVars(rw.b.color));
            if (rw.kind === "act") {
              le.setAttribute("data-act", ln.key + ":" + rw.row);
              le.innerHTML = '<i class="poap-rl-dot"></i><span class="poap-rl-name">' + esc(rw.type) + '</span><span class="poap-rl-meta">' + rw.n + " · " + Math.round(rw.pct * 100) + "%</span>";
            } else {
              var b = rw.b;
              le.setAttribute("data-id", b.id); le.setAttribute("role", "button"); le.setAttribute("tabindex", "0");
              le.innerHTML = (b.rag ? '<i class="poap-rag poap-rag-' + b.rag.toLowerCase() + '"></i>' : '<i class="poap-rag poap-rag-none"></i>') + '<span class="poap-rl-name" title="' + esc(b.item) + '">' + esc(b.item) + "</span>" +
                '<span class="poap-rl-meta">' + (b.s != null ? fmtD(b.s, true) : "TBC") + "</span>";
            }
            R.labels[lk] = le; fragL.appendChild(le);
          }
          var gk = "rg:" + ln.key + ":" + rw.row; used.rows[gk] = 1;
          if (!R.rows[gk]) {
            var ge2 = document.createElement("div"); ge2.className = "poap-rg poap-rg-" + rw.kind;
            ge2.style.cssText = "top:" + ry + "px;height:" + RM.rowH + "px;" + (rw.kind === "act" ? "--tc:" + rw.color + ";" + tintVars(rw.color) : "");
            R.rows[gk] = ge2; fragR.appendChild(ge2);
          }
          if (rw.kind === "act") {
            var ak = "as:" + ln.key + ":" + rw.row; used.items[ak] = 1;
            if (!R.items[ak]) {
              var ae = document.createElement("div"); ae.className = "poap-bsum poap-asum";
              ae.setAttribute("data-tip", rw.type + " · " + plural(rw.n, "item") + " · " + Math.round(rw.pct * 100) + "% complete · " + fmtD(rw.smin, true) + " – " + fmtD(rw.smax));
              ae.style.cssText = "left:" + ((rw.smin - minDn) * pxd) + "px;top:" + (ry + RM.rowH / 2 - 3) + "px;width:" + Math.max((rw.smax - rw.smin + 1) * pxd, 6) + "px;--pc:" + rw.color + ";" + tintVars(rw.color);
              ae.innerHTML = '<i style="width:' + Math.round(rw.pct * 100) + '%"></i>';
              R.items[ak] = ae; fragI.appendChild(ae);
            }
          }
        });
        ln.items.forEach(function (it) {
          if (it.x + it.w < x0 || it.x > x1) return;
          var bk = "b:" + it.b.id; used.items[bk] = 1;
          if (!R.items[bk]) { var be = makeBarEl(it, ln); R.items[bk] = be; fragI.appendChild(be); }
          if (it.b.bs != null && it.b.slip !== 0) {
            var gk = "g:" + it.b.id; used.ghosts[gk] = 1;
            if (!R.ghosts[gk]) { var ge = makeGhostEl(it, ln); R.ghosts[gk] = ge; fragI.appendChild(ge); }
          }
        });
      }
      // milestones
      var cy = ln.kind === "pillar" ? ln.y + ln.h / 2 : (ln.collapsed ? ln.y + ln.h / 2 : ln.y + ln.msTop / 2 - 3);
      ln.ms.forEach(function (m) {
        if (m._x + 20 < x0 || m._x - 8 > x1) return;
        var mk = "m:" + m.id; used.items[mk] = 1;
        if (!R.items[mk]) { var me = makeMsEl(m, ln, cy + (ln.kind === "pillar" ? 0 : 0)); R.items[mk] = me; fragI.appendChild(me); }
      });
    });
    // remove stale
    ["items", "rows", "labels", "ghosts"].forEach(function (grp) {
      var host = grp === "items" || grp === "ghosts" ? RM.items : (grp === "rows" ? RM.rows : RM.labels);
      var usedG = grp === "ghosts" ? used.ghosts : used[grp];
      for (var k in R[grp]) {
        if (!usedG[k]) { var e = R[grp][k]; if (e.parentNode === host) host.removeChild(e); delete R[grp][k]; }
      }
    });
    RM.rows.appendChild(fragR); RM.labels.appendChild(fragL); RM.items.appendChild(fragI);
  }

  // ---- dependency arrows -----------------------------------------
  function rmDrawDeps() {
    if (!RM.depG || !S.model) return;
    var M = S.model, ids = {};
    if (S.selKind === "bar" && S.selId && M.byId[S.selId]) ids[S.selId] = "sel";
    if (S.hoverId && M.byId[S.hoverId]) ids[S.hoverId] = ids[S.hoverId] || "hover";
    var paths = [], linked = {};
    Object.keys(ids).forEach(function (id) {
      var b = M.byId[id];
      b.deps.forEach(function (pid) { paths.push([pid, id, ids[id]]); linked[pid] = 1; });
      b.succ.forEach(function (sid) { paths.push([id, sid, ids[id]]); linked[sid] = 1; });
      linked[id] = 1;
    });
    var out = "", seen = {};
    paths.forEach(function (p) {
      var k = p[0] + ">" + p[1]; if (seen[k]) return; seen[k] = 1;
      var a = RM.geo[p[0]], z = RM.geo[p[1]];
      if (!a || !z) return;
      var x1 = a.x + a.w, y1 = a.cy, x2 = z.x, y2 = z.cy;
      var dx = Math.max(24, Math.abs(x2 - x1) / 2);
      var d;
      if (x2 >= x1 + 6) d = "M" + x1 + " " + y1 + " C" + (x1 + dx) + " " + y1 + " " + (x2 - dx) + " " + y2 + " " + x2 + " " + y2;
      else { var mid = Math.min(y1, y2) - 0; d = "M" + x1 + " " + y1 + " C" + (x1 + 40) + " " + y1 + " " + (x2 - 40) + " " + y2 + " " + x2 + " " + y2; }
      var isSel = p[2] === "sel";
      out += '<path d="' + d + '" fill="none" stroke="' + (isSel ? "#0f1c3f" : "#0ea5a0") + '" stroke-width="' + (isSel ? 2 : 1.8) + '" stroke-linecap="round" ' + (isSel ? "" : 'stroke-dasharray="0"') + ' marker-end="url(#' + (isSel ? "poap-ah" : "poap-ah2") + ')" opacity=".9"/>';
    });
    RM.depG.innerHTML = out;
    // link highlight on rendered bars
    for (var k2 in RM.rendered.items) {
      if (k2.charAt(0) !== "b") continue;
      var e = RM.rendered.items[k2], bid = e.getAttribute("data-id");
      cls(e, "is-link", !!linked[bid] && !ids[bid]);
      cls(e, "is-dim", Object.keys(ids).length > 0 && !linked[bid]);
    }
  }

  // ---- minimap ----------------------------------------------------
  function rmDrawMinimap() {
    var M = S.model, c = RM.mmCanvas;
    var W = RM.mm.clientWidth, H = RM.mm.clientHeight || 56;
    if (!W) return;
    var dpr = window.devicePixelRatio || 1;
    c.width = Math.round(W * dpr); c.height = Math.round(H * dpr); c.style.width = W + "px"; c.style.height = H + "px";
    var ctx = c.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, W, H);
    var span = M.maxDn - M.minDn + 1, sx = W / span;
    var top = 4, bottom = H - 15;
    // month shading + labels
    var d = dnParts(M.minDn), cur = dnFromYMD(d.y, d.m + 1, 1), i = 0;
    ctx.font = "600 9.5px -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif";
    ctx.textBaseline = "alphabetic";
    while (cur <= M.maxDn) {
      var p = dnParts(cur), nx = dnFromYMD(p.m === 11 ? p.y + 1 : p.y, p.m === 11 ? 1 : p.m + 2, 1);
      var x0 = (cur - M.minDn) * sx, x1 = (nx - M.minDn) * sx;
      if (i % 2 === 0) { ctx.fillStyle = "rgba(15,28,63,.035)"; ctx.fillRect(x0, 0, x1 - x0, H); }
      if (p.m === 0 || p.m === 3 || p.m === 6 || p.m === 9) {
        ctx.fillStyle = "rgba(15,28,63,.12)"; ctx.fillRect(x0, 0, 1, H);
        ctx.fillStyle = p.m === 0 ? "#344054" : "#98a2b3";
        ctx.fillText(p.m === 0 ? String(p.y) : "Q" + (p.m / 3 + 1), x0 + 4, H - 4);
      }
      cur = nx; i++;
    }
    // lanes
    var lanes = RM.layout.lanes.filter(function (l) { return l.kind === "topic"; });
    var lh = Math.max(1.4, Math.min(5, (bottom - top) / Math.max(1, lanes.length)));
    lanes.forEach(function (ln, li) {
      var y = top + li * lh;
      ln.bars.forEach(function (b) {
        var x = (b.ps - M.minDn) * sx, w = Math.max((b.pe - b.ps + 1) * sx, 1);
        ctx.globalAlpha = b.tbc ? 0.35 : 0.9; ctx.fillStyle = b.color; ctx.fillRect(x, y, w, Math.max(1, lh - 0.5));
      });
    });
    ctx.globalAlpha = 1;
    // milestones as ticks at the top
    S.fm.forEach(function (m) { ctx.fillStyle = "rgba(15,28,63,.55)"; ctx.fillRect((m.dn - M.minDn) * sx - 0.5, 0, 1, 3); });
    // today
    var tx = (M.todayDn - M.minDn) * sx;
    ctx.fillStyle = "#ef4444"; ctx.fillRect(tx - 0.5, 0, 1.4, H - 12);
    if (M.statusDn !== M.todayDn) { ctx.fillStyle = "#0f1c3f"; ctx.fillRect((M.statusDn - M.minDn) * sx - 0.5, 0, 1, H - 12); }
  }
  function rmUpdateWin() {
    if (!RM.layout) return;
    var W = RM.mm.clientWidth, tl = RM.layout.tlW; if (!W || !tl) return;
    var sc = RM.sc, vis = Math.max(20, sc.clientWidth - RM.labelW);
    var left = sc.scrollLeft / tl * W, w = clamp(vis / tl * W, 12, W);
    left = clamp(left, 0, W - w);
    RM.mmWin.style.left = left + "px"; RM.mmWin.style.width = w + "px";
    RM.mmWin.setAttribute("aria-valuenow", String(Math.round(sc.scrollLeft / Math.max(1, tl - vis) * 100)));
    RM.mmWin.setAttribute("aria-valuemin", "0"); RM.mmWin.setAttribute("aria-valuemax", "100");
  }
  function rmMmLeft(frac) {
    var W = RM.mm.clientWidth, tl = RM.layout.tlW;
    var winW = RM.mmWin.offsetWidth;
    var left = clamp(frac * W, 0, W - winW);
    RM.sc.scrollLeft = left / W * tl;
  }
  function rmMmCenter(frac) {
    var W = RM.mm.clientWidth, winW = RM.mmWin.offsetWidth;
    rmMmLeft(frac - winW / 2 / W);
  }

  // ---- zoom / scroll actions -------------------------------------
  function rmCenterDn() {
    var sc = RM.sc; return S.model.minDn + (sc.scrollLeft + (sc.clientWidth - RM.labelW) / 2) / S.pxd;
  }
  function rmSetPxd(pxd, key, centerDn) {
    var sc = RM.sc;
    var cdn = centerDn == null ? rmCenterDn() : centerDn;
    S.pxd = pxd; S.zoom = key;
    rmLayoutAll();
    sc.scrollLeft = Math.max(0, (cdn - S.model.minDn) * S.pxd - (sc.clientWidth - RM.labelW) / 2);
    rmUpdateWin();
  }
  function rmSetZoom(k) { if (!ZOOMS[k]) return; rmSetPxd(ZOOMS[k], k); }
  function rmZoomTo(pxd, clientX) {
    var sc = RM.sc, r = sc.getBoundingClientRect();
    var tx = sc.scrollLeft + (clientX - r.left) - RM.labelW;
    var dn = S.model.minDn + tx / S.pxd;
    var oldOff = (clientX - r.left);
    S.pxd = pxd; S.zoom = "custom";
    rmLayoutAll();
    sc.scrollLeft = Math.max(0, (dn - S.model.minDn) * pxd + RM.labelW - oldOff);
    rmUpdateWin();
  }
  function rmFit() {
    var sc = RM.sc, M = S.model;
    var avail = Math.max(300, sc.clientWidth - RM.labelW - 20);
    var pxd = clamp(avail / (M.maxDn - M.minDn + 1), 0.12, 14);
    S.pxd = pxd; S.zoom = "fit";
    rmLayoutAll(); sc.scrollLeft = 0; rmUpdateWin();
  }
  // PDF export: draw the whole roadmap (not just the scrolled window), fitted to the page width
  ATSpoap.prepareExport = function () {
    if (!S.model || S.view !== "roadmap" || !RM.ready) return null;
    var saved = { zoom: S.zoom, pxd: S.pxd, sl: RM.sc.scrollLeft, st: RM.sc.scrollTop };
    RM.sc.style.height = "auto"; RM.sc.style.overflow = "visible";
    RM.printing = true; rmHl(null); rmFit();
    return {
      restore: function () {
        RM.sc.style.height = ""; RM.sc.style.overflow = ""; RM.printing = false;
        S.zoom = saved.zoom; S.pxd = saved.pxd; rmLayoutAll();
        RM.sc.scrollLeft = saved.sl; RM.sc.scrollTop = saved.st; rmUpdateWin(); RM.zoomSeg.setActive(S.zoom); cls(RM.btnFit, "is-on", S.zoom === "fit");
      }
    };
  };
  function rmToday(smooth) {
    var sc = RM.sc, M = S.model;
    var x = (M.todayDn - M.minDn) * S.pxd;
    var target = Math.max(0, x - (sc.clientWidth - RM.labelW) * 0.33);
    try { sc.scrollTo({ left: target, top: sc.scrollTop, behavior: smooth ? "smooth" : "auto" }); } catch (e) { sc.scrollLeft = target; }
  }
  function rmToggle(key) {
    S.collapsed[key] = !S.collapsed[key];
    var sc = RM.sc, st = sc.scrollTop, sl = sc.scrollLeft;
    rmLayoutAll(true);
    sc.scrollTop = st; sc.scrollLeft = sl;
  }
  function rmSetAll(collapse) {
    var M = S.model;
    if (collapse) { RM.layout.lanes.forEach(function (l) { if (l.kind === "topic") S.collapsed[l.key] = true; }); }
    else { S.collapsed = {}; }
    var sc = RM.sc, st = sc.scrollTop, sl = sc.scrollLeft;
    rmLayoutAll(true); sc.scrollTop = st; sc.scrollLeft = sl;
  }
  function rmBuildLegend() {
    var M = S.model, present = {};
    M.bars.forEach(function (b) { present[b.type] = 1; });
    var types = M.types.filter(function (t) { return present[t.name] && t.name !== "Holiday / Freeze"; });
    var html = '<div class="poap-leg-grp"><span class="poap-leg-h">Activity</span>' + types.map(function (t) {
      return '<button type="button" class="poap-leg-i' + (S.filters.type === t.name ? " on" : "") + '" data-type="' + esc(t.name) + '" title="Filter to ' + esc(t.name) + '"><i class="poap-sw" style="background:' + t.color + '"></i>' + esc(t.name) + "</button>";
    }).join("") + (present["Holiday / Freeze"] ? '<span class="poap-leg-i"><i class="poap-sw poap-sw-hatch"></i>Holiday / freeze</span>' : "") + "</div>";
    html += '<div class="poap-leg-grp"><span class="poap-leg-h">RAG</span><span class="poap-leg-i"><i class="poap-rag poap-rag-green"></i>Green</span><span class="poap-leg-i"><i class="poap-rag poap-rag-amber"></i>Amber</span><span class="poap-leg-i"><i class="poap-rag poap-rag-red"></i>Red</span></div>';
    html += '<div class="poap-leg-grp"><span class="poap-leg-h">Marks</span>' +
      '<span class="poap-leg-i"><i class="poap-leg-ghost"></i>Baseline</span><span class="poap-leg-i"><b class="poap-bar-slip is-lg">+14d</b>Slip vs baseline</span>' +
      '<span class="poap-leg-i"><i class="poap-dia" style="background:#0f1c3f"></i>Milestone</span><span class="poap-leg-i"><i class="poap-leg-tbc"></i>Dates TBC</span>' +
      '<span class="poap-leg-i"><i class="poap-leg-today"></i>Today</span><span class="poap-leg-i"><i class="poap-leg-fill"></i>% complete</span></div>';
    RM.legend.innerHTML = html;
    Array.prototype.forEach.call(RM.legend.querySelectorAll("[data-type]"), function (b) {
      b.addEventListener("click", function () { var t = b.getAttribute("data-type"); setFilter({ type: S.filters.type === t ? "" : t }); });
    });
  }

  // ---- public-ish: focus -----------------------------------------
  function focusBar(id, opts) {
    opts = opts || {};
    var M = S.model; if (!M) return false;
    var b = M.byId[id];
    if (!b) return false;
    if (S.view !== "roadmap") setView("roadmap");
    if (!S.fb.some(function (x) { return x.id === id; })) { resetFilters(true); }
    delete S.collapsed["p:" + b.pillar]; delete S.collapsed["t:" + b.key];
    S.views.roadmap.render();
    var g = RM.geo[id]; if (!g) return false;
    var sc = RM.sc;
    var ln = g.lane;
    var cy = g.cy + RM.headH;
    var left = Math.max(0, g.x + g.w / 2 - (sc.clientWidth - RM.labelW) / 2);
    var top = Math.max(0, cy - sc.clientHeight / 2);
    try { sc.scrollTo({ left: left, top: top, behavior: opts.instant ? "auto" : "smooth" }); } catch (e) { sc.scrollLeft = left; sc.scrollTop = top; }
    rmDraw();
    if (opts.drawer !== false) openDrawer("bar", id);
    // flash after scroll settles
    setTimeout(function () {
      rmDraw();
      var el = RM.rendered.items["b:" + id];
      if (el) { el.classList.remove("is-flash"); void el.offsetWidth; el.classList.add("is-flash"); setTimeout(function () { el.classList.remove("is-flash"); }, 1800); }
    }, opts.instant ? 20 : 450);
    return true;
  }
  function focusMs(id) {
    var M = S.model; if (!M || !M.msById[id]) return false;
    var m = M.msById[id];
    if (S.view !== "roadmap") setView("roadmap");
    if (!S.fm.some(function (x) { return x.id === id; })) resetFilters(true);
    delete S.collapsed["p:" + (m.pillar || "Cross-cutting")];
    S.views.roadmap.render();
    var sc = RM.sc;
    var left = Math.max(0, (m.dn - M.minDn) * S.pxd - (sc.clientWidth - RM.labelW) / 2);
    var ln = RM.layout.lanes.filter(function (l) { return l.ms.indexOf(m) >= 0; })[0];
    var top = ln ? Math.max(0, ln.y + RM.headH - sc.clientHeight / 2) : sc.scrollTop;
    try { sc.scrollTo({ left: left, top: top, behavior: "smooth" }); } catch (e) { sc.scrollLeft = left; sc.scrollTop = top; }
    openDrawer("ms", id);
    return true;
  }
  function showOnRoadmap(ref) {
    if (ref.indexOf("ms:") === 0) return focusMs(ref.slice(3));
    return focusBar(ref);
  }

  function rmInitialScroll() {
    if (RM.inited || RM.sc.clientWidth < 400) return;
    RM.inited = true; RM.sc.scrollTop = 0; RM.sc.scrollLeft = 0; rmToday(false);
    rmDraw(); rmUpdateWin();
  }
  S.views.roadmap = {
    mount: rmMount,
    render: function () {
      if (!RM.ready || !S.model) return;
      rmBuildLegend();
      rmLayoutAll();
      rmInitialScroll();
      S.dirty.roadmap = false;
    },
    resize: function () {
      if (!RM.ready || !S.model) return;
      rmInitialScroll();
      if (S.zoom === "fit") { rmFit(); return; }
      rmDraw(); rmDrawMinimap(); rmUpdateWin();
    },
    onSelect: function () { if (RM.ready) { rmDrawDeps(); } }
  };

  // ===============================================================
  // PLAN ON A PAGE (executive poster)
  // ===============================================================
  var PO = { ready: false, lists: { next: false, risk: false } };

  function poMount(container) {
    var el = h("div", { class: "poap-po" });
    var tb = h("div", { class: "poap-rm-tb" });
    PO.densSeg = segmented([{ key: "comfortable", label: "Comfortable", title: "One row per topic" }, { key: "compact", label: "Compact", title: "Tighter rows" }, { key: "summary", label: "Pillars", title: "One-screen summary: one row per pillar" }], S.dens, function (k) { S.dens = k; S.views.poap.render(); }, "Row density");
    var grpA = h("div", { class: "poap-tb-grp" }, [PO.densSeg]);
    PO.count = h("div", { class: "poap-rm-count" });
    var pr = h("button", { type: "button", class: "btn btn-outline btn-sm poap-noprint", title: "Save the one-page poster as a PDF", html: ICON.pdf + "Export PDF" });
    pr.addEventListener("click", function () { if (window.ATS && ATS.exportPdf) ATS.exportPdf(); });
    tb.appendChild(grpA); tb.appendChild(h("div", { class: "poap-tb-grp" }, [pr])); tb.appendChild(PO.count);
    el.appendChild(tb);
    var grid = h("div", { class: "poap-po-grid" });
    PO.main = h("div", { class: "poap-po-main card" });
    PO.side = h("div", { class: "poap-po-side" });
    grid.appendChild(PO.main); grid.appendChild(PO.side);
    el.appendChild(grid);
    container.appendChild(el);
    PO.el = el;
    PO.main.addEventListener("click", function (ev) {
      var pc2 = ev.target.closest("[data-pcol]");
      if (pc2) { var k = "po:" + pc2.getAttribute("data-pcol"); S.collapsed[k] = !S.collapsed[k]; poRender(); return; }
      var c = ev.target.closest("[data-gid]");
      if (c) { var g = PO.groups[c.getAttribute("data-gid")]; if (g) poOpenGroup(g); return; }
      var m = ev.target.closest("[data-msid]"); if (m) { openDrawer("ms", m.getAttribute("data-msid")); return; }
    });
    PO.main.addEventListener("mouseover", function (ev) {
      var c = ev.target.closest("[data-gid]");
      if (c) { var g = PO.groups[c.getAttribute("data-gid")]; if (g) showTip(groupTipHtml(g), ev); return; }
      var m = ev.target.closest("[data-msid]"); if (m) { showTip(msTipHtml(S.model.msById[m.getAttribute("data-msid")]), ev); return; }
      var t = ev.target.closest("[data-tip]"); if (t) showTip('<div class="poap-tip-t">' + esc(t.getAttribute("data-tip")) + "</div>", ev);
    });
    PO.main.addEventListener("keydown", function (ev) {
      if (ev.key !== "Enter" && ev.key !== " ") return;
      var pc2 = ev.target.closest("[data-pcol]");
      if (pc2) { ev.preventDefault(); var k = "po:" + pc2.getAttribute("data-pcol"); S.collapsed[k] = !S.collapsed[k]; poRender(); }
    });
    PO.main.addEventListener("mousemove", moveTip);
    PO.main.addEventListener("mouseout", function (ev) {
      if (!ev.relatedTarget || !ev.relatedTarget.closest || !ev.relatedTarget.closest("[data-gid],[data-msid],[data-tip]")) hideTip();
    });
    PO.side.addEventListener("click", function (ev) {
      var t = ev.target.closest("[data-poap-open]");
      if (t) { openDrawer(t.getAttribute("data-poap-kind") || "bar", t.getAttribute("data-poap-open")); return; }
      var x = ev.target.closest("[data-more]");
      if (x) { var k = x.getAttribute("data-more"); PO.lists[k] = !PO.lists[k]; poSide(); }
    });
    PO.ready = true;
  }

  function aggStatus(bars) {
    var anyLate = bars.some(function (b) { return b.status === "Delayed" || b.slip > 0; });
    if (anyLate) return "Delayed";
    if (bars.every(function (b) { return b.status === "Complete"; })) return "Complete";
    if (bars.some(function (b) { return b.status === "In Progress"; }) || (bars.some(function (b) { return b.status === "Complete"; }) && bars.some(function (b) { return b.status !== "Complete"; }))) return "In Progress";
    if (bars.every(function (b) { return b.tbc; })) return "Dates TBC";
    if (bars.some(function (b) { return b.status === "On Hold"; })) return "On Hold";
    return "Not Started";
  }
  function poClusters(bars) {
    var out = [];
    PHASES.forEach(function (ph, pi) {
      var list = bars.filter(function (b) { return ph.types.indexOf(b.type) >= 0; }).sort(function (a, b) { return a.ps - b.ps; });
      var cur = null;
      list.forEach(function (b) {
        if (cur && b.ps <= cur.e + 21) { cur.bars.push(b); cur.e = Math.max(cur.e, b.pe); }
        else { cur = { phase: ph, pi: pi, bars: [b], s: b.ps, e: b.pe }; out.push(cur); }
      });
    });
    out.forEach(function (c) {
      c.status = aggStatus(c.bars); c.pct = weightedPct(c.bars); c.rag = worstRag(c.bars);
      c.tbc = c.bars.every(function (b) { return b.indicative; });
      c.slip = Math.max.apply(null, c.bars.map(function (b) { return b.slip; }));
      c.realS = Math.min.apply(null, c.bars.map(function (b) { return b.s != null ? b.s : 1e9; }));
      c.realE = Math.max.apply(null, c.bars.map(function (b) { return b.e != null ? b.e : -1e9; }));
      c.color = c.bars[0].color;
    });
    return out;
  }
  function groupTipHtml(g) {
    return '<div class="poap-tip-h"><i class="poap-sw" style="background:' + g.color + '"></i>' + esc(g.phase.label) + "<span>" + plural(g.bars.length, "activity").replace("activitys", "activities") + "</span></div>" +
      '<div class="poap-tip-t">' + esc(g.topic) + "</div><div class=\"poap-tip-s\">" + esc(g.pillar) + "</div>" +
      '<div class="poap-tip-r">' + (g.realS < 1e8 ? fmtD(g.realS) + " → " + fmtD(g.realE) : "Dates TBC (indicative position)") + "</div>" +
      '<div class="poap-tip-r">' + chipHtml(g.status, statusKind(g.status)) + " " + ragDot(g.rag) + " " + Math.round(g.pct * 100) + "% complete" + (g.slip > 0 ? ' <b class="poap-neg">+' + g.slip + "d slip</b>" : "") + "</div>";
  }
  function poOpenGroup(g) {
    if (g.bars.length === 1) { openDrawer("bar", g.bars[0].id); return; }
    var id = g.gid; groupStore[id] = { phase: g.phase.label, topic: g.topic, pillar: g.pillar, color: g.color, status: g.status, rag: g.rag, pct: g.pct, s: g.realS < 1e8 ? g.realS : null, e: g.realE, bars: g.bars };
    openDrawer("group", id);
  }

  function poRange() {
    var M = S.model, all = [M.statusDn, M.todayDn];
    S.fb.forEach(function (b) { all.push(b.ps, b.pe); });
    S.fm.forEach(function (m) { if (m.key2 === "key") all.push(m.dn); });
    var mn = Math.min.apply(null, all), mx = Math.max.apply(null, all);
    var a = dnParts(mn), z = dnParts(mx);
    var qs = dnFromYMD(a.y, Math.floor(a.m / 3) * 3 + 1, 1);
    var qm = Math.floor(z.m / 3) * 3 + 4, qy = z.y; if (qm > 12) { qm -= 12; qy++; }
    return { min: qs, max: dnFromYMD(qy, qm, 1) - 1 };
  }
  function pct100(v) { return (Math.round(v * 1000) / 10) + "%"; }

  function poRender() {
    if (!PO.ready || !S.model) return;
    var M = S.model;
    PO.densSeg.setActive(S.dens);
    PO.el.classList.toggle("is-compact", S.dens !== "comfortable");
    var bars = S.fb.filter(function (b) { return phaseOf(b.type) >= 0; });
    if (!bars.length) {
      PO.main.innerHTML = emptyState(M.bars.length ? "Nothing to show for these filters" : "No activities in this workbook", M.bars.length ? "Clear a filter, or choose an activity type that is part of the test phases." : "Add rows to the POAP_Plan sheet and reload the workbook.");
      PO.side.innerHTML = ""; PO.count.innerHTML = ""; PO.groups = {};
      S.dirty.poap = false; return;
    }
    var rg = poRange(), span = rg.max - rg.min + 1;
    function pc(dn) { return ((dn - rg.min) / span * 100); }
    var labelW = 212;
    var tlPx = Math.max(480, (PO.main.clientWidth || 900) - labelW - 24);
    PO.tlPx = tlPx;
    var comp = S.dens !== "comfortable";
    var chH = comp ? 22 : 26, subH = comp ? 26 : 31;
    PO.groups = {};
    var gcount = 0;

    // header: years + quarters
    var qhtml = "", yhtml = "", vlines = "";
    var d = dnParts(rg.min), cy = d.y, cq = Math.floor(d.m / 3), cur = rg.min, lastY = null;
    while (cur <= rg.max) {
      var p = dnParts(cur), q = Math.floor(p.m / 3);
      var nm = (q === 3) ? 1 : (q + 1) * 3 + 1, ny = (q === 3) ? p.y + 1 : p.y;
      var nxt = dnFromYMD(ny, nm, 1);
      var left = pc(cur), w = (Math.min(nxt, rg.max + 1) - cur) / span * 100;
      qhtml += '<div class="poap-qc" style="left:' + left + "%;width:" + w + '%"><span>Q' + (q + 1) + (w / 100 * tlPx > 70 ? " " + p.y : "") + "</span></div>";
      vlines += '<i class="poap-qline" style="left:' + left + '%"></i>';
      if (lastY !== p.y) {
        var yEnd = dnFromYMD(p.y + 1, 1, 1);
        yhtml += '<div class="poap-yc" style="left:' + left + "%;width:" + ((Math.min(yEnd, rg.max + 1) - cur) / span * 100) + '%"><span>' + p.y + "</span></div>"; lastY = p.y;
      }
      cur = nxt;
    }
    var todayX = pc(M.todayDn), statusX = pc(M.statusDn);
    var lines = '<i class="poap-po-today" style="left:' + todayX + '%"></i>';
    if (M.statusDn !== M.todayDn) lines += '<i class="poap-po-status" style="left:' + statusX + '%"></i>';

    // milestone strip (key milestones)
    var keyMs = S.fm.filter(function (m) { return m.key2 === "key"; }).sort(function (a, b) { return a.dn - b.dn; });
    var stripLast = [-1e9, -1e9, -1e9], msHtml = "";
    keyMs.forEach(function (m) {
      var x = pc(m.dn) / 100 * tlPx, wantLbl = (m.type !== "Sign-off / TCR");
      var lw = Math.min(150, m.name.length * 5.4 + 22), lane = -1, labelled = false, i;
      if (wantLbl) for (i = 0; i < 3; i++) if (x - stripLast[i] > 6) { lane = i; labelled = true; break; }
      if (lane < 0) for (i = 0; i < 3; i++) if (x - stripLast[i] > 3) { lane = i; break; }
      if (lane < 0) lane = stripLast.indexOf(Math.min.apply(null, stripLast));
      stripLast[lane] = x + (labelled ? lw : 13);
      msHtml += '<button type="button" class="poap-pms' + (labelled ? " has-lbl" : "") + '" data-msid="' + esc(m.id) + '" style="left:' + pc(m.dn) + "%;top:" + (1 + lane * 16) + 'px;--mc:' + m.color + '" aria-label="Milestone ' + esc(m.name) + " " + fmtD(m.dn) + '"><i class="poap-dia"></i>' +
        (labelled ? "<span>" + esc(m.name) + "</span>" : "") + "</button>";
    });

    var html = '<div class="poap-poster">';
    html += '<div class="poap-po-title"><div><div class="poap-po-h1">' + esc(M.program) + '</div><div class="poap-po-h2">Plan on a Page · test delivery roadmap · as at ' + fmtD(M.statusDn) + "</div></div>" +
      '<div class="poap-po-key">' + ["Complete", "In Progress", "Not Started", "Delayed", "Dates TBC"].map(function (s) { return '<span class="poap-keyi"><i class="poap-chv-sw st-' + stKey(s) + '"></i>' + s + "</span>"; }).join("") + "</div></div>";
    html += '<div class="poap-po-head" style="--lw:' + labelW + 'px"><div class="poap-po-hl"><span>Pillar › Topic</span></div><div class="poap-po-ht"><div class="poap-yrow">' + yhtml + '</div><div class="poap-qrow">' + qhtml + '</div><div class="poap-mstrip"><span class="poap-mstrip-l">Key milestones</span>' + msHtml + '</div><div class="poap-hlines">' + lines + "</div></div></div>";
    html += '<div class="poap-po-rows" style="--lw:' + labelW + 'px">';
    // pillar groups
    var byP = {};
    bars.forEach(function (b) { (byP[b.pillar] = byP[b.pillar] || { topics: {}, order: [], bars: [] }); var P = byP[b.pillar]; if (!P.topics[b.topic]) { P.topics[b.topic] = { bars: [], first: b.idx }; P.order.push(b.topic); } P.topics[b.topic].bars.push(b); P.bars.push(b); });
    var pnames = Object.keys(byP).sort(function (a, b) { return (M.pillarIdx[a] == null ? 99 : M.pillarIdx[a]) - (M.pillarIdx[b] == null ? 99 : M.pillarIdx[b]); });
    var minW = 26 / tlPx * 100, nest = 9 / tlPx * 100;
    // pack clusters into sub-rows and emit chevrons; returns { html, nsub }
    function chevrons(cl, tn, pn) {
      var rowsEnd = [], out = "";
      cl.sort(function (a, b) { return a.s - b.s || a.pi - b.pi; }).forEach(function (c) {
        c.left = pc(c.s); c.w = Math.max((c.e - c.s + 1) / span * 100, minW);
        var r = 0; while (r < rowsEnd.length && rowsEnd[r] > c.left + nest) r++;
        rowsEnd[r] = c.left + c.w; c.r = r;
      });
      cl.forEach(function (c) {
        var gid = "g" + (gcount++);
        c.gid = gid; c.topic = tn; c.pillar = pn; PO.groups[gid] = c;
        var pxw = c.w / 100 * tlPx;
        var dates = c.realS < 1e8 ? fmtShort(c.realS) + "–" + fmtShort(c.realE) : "TBC";
        var txt = pxw >= 38 ? '<span class="poap-chv-t">' + esc(pxw >= 72 ? c.phase.label : c.phase.short) + "</span>" : "";
        if (pxw >= 150) txt += '<span class="poap-chv-d">' + esc(dates) + "</span>";
        var dot = (c.rag && pxw >= 54) ? '<i class="poap-rag poap-rag-' + c.rag.toLowerCase() + '"></i>' : "";
        var cnt = c.bars.length > 1 && pxw >= 90 ? '<em>×' + c.bars.length + "</em>" : "";
        out += '<div class="poap-chv-w" style="left:' + c.left + "%;width:" + c.w + "%;top:" + (5 + c.r * subH) + "px;height:" + chH + 'px"><button type="button" class="poap-chv st-' + stKey(c.status) + '" data-gid="' + gid + '" data-sel="bar:' + (c.bars.length === 1 ? c.bars[0].id : "") + '" style="--p:' + Math.round(c.pct * 100) + '%;height:' + chH + 'px" aria-label="' + esc(tn + ": " + c.phase.label + ", " + c.status + ", " + dates) + '">' + dot + txt + cnt + (c.slip > 0 && pxw >= 100 ? '<b class="poap-chv-slip">+' + c.slip + "d</b>" : "") + "</button></div>";
      });
      return { html: out, nsub: Math.max(1, rowsEnd.length) };
    }
    var nTopics = 0;
    pnames.forEach(function (pn) {
      var P = byP[pn], pcol = M.pillarColor[pn] || "#64748b";
      var pp = weightedPct(P.bars), smin = Math.min.apply(null, P.bars.map(function (b) { return b.ps; })), smax = Math.max.apply(null, P.bars.map(function (b) { return b.pe; }));
      var asSummary = S.dens === "summary" || !!S.collapsed["po:" + pn];
      var pcl = null, pch = null, phtml = "", ph = 38;
      if (asSummary) {
        pcl = poClusters(P.bars); pch = chevrons(pcl, pn + " · all topics", pn);
        ph = pch.nsub * subH + 12; phtml = pch.html;
      } else {
        phtml = '<div class="poap-po-psum" data-tip="' + esc(pn + " · " + plural(P.bars.length, "activity").replace("activitys", "activities") + " · " + pct100(pp) + " complete") + '" style="left:' + pc(smin) + "%;width:" + Math.max(0.4, (smax - smin + 1) / span * 100) + '%"><i style="width:' + Math.round(pp * 100) + '%"></i></div>';
      }
      var coll = !!S.collapsed["po:" + pn];
      html += '<div class="poap-po-prow' + (asSummary ? " is-sum" : "") + '" style="--pc:' + pcol + ";height:" + ph + 'px"><div class="poap-po-pl"' + (S.dens === "summary" ? "" : ' role="button" tabindex="0" data-pcol="' + esc(pn) + '" aria-expanded="' + (coll ? "false" : "true") + '" aria-label="' + (coll ? "Expand " : "Collapse ") + esc(pn) + '"') + ">" + (S.dens === "summary" ? "" : '<span class="poap-chev' + (coll ? " is-c" : "") + '">' + ICON.chev + "</span>") + lanePctRing(pp, pcol) + '<b>' + esc(pn) + '</b><span>' + P.order.length + " topics · " + pct100(pp) + '</span></div><div class="poap-po-pt">' + vlines + phtml + lines + "</div></div>";
      nTopics += P.order.length;
      if (asSummary) return;
      var tnames = P.order.slice().sort(function (a, b) { return P.topics[a].first - P.topics[b].first; });
      tnames.forEach(function (tn, ti) {
        var T = P.topics[tn], ch = chevrons(poClusters(T.bars), tn, pn);
        var rowH = ch.nsub * subH + 10;
        var tl = weightedPct(T.bars), rag = worstRag(T.bars);
        html += '<div class="poap-po-trow' + (ti % 2 ? " odd" : "") + '" style="height:' + rowH + 'px"><div class="poap-po-tl"><span class="poap-po-tn" title="' + esc(tn) + '">' + esc(tn) + '</span><span class="poap-po-ts">' + pct100(tl) + (rag ? " " + ragDot(rag) : "") + '</span></div><div class="poap-po-tt">' + vlines + ch.html + lines + "</div></div>";
      });
    });
    html += "</div></div>";
    PO.main.innerHTML = html;
    PO.count.innerHTML = "<b>" + bars.length + "</b> activities across <b>" + nTopics + "</b> topics in <b>" + pnames.length + "</b> pillars";
    poSide();
    markSelected();
    S.dirty.poap = false;
  }
  function stKey(s) { return { "Complete": "done", "In Progress": "prog", "Not Started": "todo", "Delayed": "late", "Dates TBC": "tbc", "On Hold": "hold" }[s] || "todo"; }

  function poSide() {
    var M = S.model, st = M.statusDn, end = st + 30;
    var items = [];
    S.fb.forEach(function (b) {
      if (b.s == null || b.isHol) return;
      var starts = b.s >= st && b.s <= end, ends = b.e >= st && b.e <= end && b.s <= end;
      if (b.status === "Complete") return;
      if (starts && ends) items.push({ dn: b.s, b: b, label: "Starts " + fmtShort(b.s) + " · ends " + fmtShort(b.e) });
      else if (starts) items.push({ dn: b.s, b: b, label: "Starts " + fmtShort(b.s) + " (" + relDays(b.s) + ")" });
      else if (ends) items.push({ dn: b.e, b: b, label: "Ends " + fmtShort(b.e) + " (" + relDays(b.e) + ")" });
    });
    S.fm.forEach(function (m) { if (m.dn >= st && m.dn <= end) items.push({ dn: m.dn, m: m, label: fmtShort(m.dn) + " (" + relDays(m.dn) + ")" }); });
    items.sort(function (a, b) { return a.dn - b.dn; });
    var limit = PO.lists.next ? 200 : 8;
    var h1 = '<div class="poap-sidecard card"><div class="poap-sidehead"><div><div class="card-title">Next 30 days</div><div class="card-sub">' + fmtD(st, true) + " – " + fmtD(end) + '</div></div><span class="poap-count">' + items.length + "</span></div><div class=\"poap-sidelist\">";
    if (!items.length) h1 += '<div class="poap-sideempty">Nothing starting or ending in the next 30 days.</div>';
    items.slice(0, limit).forEach(function (it) {
      if (it.b) h1 += sideItem(it.b, it.label, "bar");
      else h1 += '<a href="#" class="poap-sitem" data-poap-open="' + esc(it.m.id) + '" data-poap-kind="ms"><i class="poap-dia" style="background:' + it.m.color + '"></i><span class="poap-si-t">' + esc(it.m.name) + '</span><span class="poap-si-m">Milestone · ' + esc(it.label) + "</span></a>";
    });
    h1 += "</div>" + (items.length > 8 ? '<button type="button" class="poap-more" data-more="next">' + (PO.lists.next ? "Show fewer" : "Show all " + items.length) + "</button>" : "") + "</div>";
    // at risk
    var risk = [];
    S.fb.forEach(function (b) {
      var sc = 0, why = [];
      if (b.status === "Delayed") { sc += 3; why.push(["Delayed", "red"]); }
      if (b.slip > 0) { sc += b.slip > 14 ? 3 : 2; why.push(["+" + b.slip + "d", b.slip > 14 ? "red" : "amber"]); }
      if (b.rag === "Red") { sc += 3; why.push(["Red", "red"]); } else if (b.rag === "Amber") { sc += 2; why.push(["Amber", "amber"]); }
      if (b.status === "Dates TBC") { sc += 1; why.push(["Dates TBC", "dash"]); }
      if (b.status === "On Hold") { sc += 1; why.push(["On hold", "amber"]); }
      if (sc) risk.push({ b: b, sc: sc, why: why });
    });
    risk.sort(function (a, b) { return b.sc - a.sc || (a.b.ps - b.b.ps); });
    var lim2 = PO.lists.risk ? 300 : 8;
    var h2 = '<div class="poap-sidecard card"><div class="poap-sidehead"><div><div class="card-title">At risk / slipping</div><div class="card-sub">Delayed, slip &gt; 0, Red/Amber, Dates TBC</div></div><span class="poap-count' + (risk.length ? " warn" : "") + '">' + risk.length + "</span></div><div class=\"poap-sidelist\">";
    if (!risk.length) h2 += '<div class="poap-sideempty">No activities at risk. 🎉</div>'.replace(" 🎉", "");
    risk.slice(0, lim2).forEach(function (r) {
      h2 += sideItem(r.b, r.why.map(function (w) { return chipHtml(w[0], w[1]); }).join(" "), "bar", true);
    });
    h2 += "</div>" + (risk.length > 8 ? '<button type="button" class="poap-more" data-more="risk">' + (PO.lists.risk ? "Show fewer" : "Show all " + risk.length) + "</button>" : "") + "</div>";
    PO.side.innerHTML = h1 + h2;
  }
  function sideItem(b, meta, kind, metaIsHtml) {
    return '<a href="#" class="poap-sitem" data-poap-open="' + esc(b.id) + '" data-poap-kind="bar"><i class="poap-sw" style="background:' + b.color + '"></i><span class="poap-si-t">' + esc(b.item) + '</span><span class="poap-si-s">' + esc(b.topic) + " · " + esc(b.type) + '</span><span class="poap-si-m">' + (metaIsHtml ? meta : esc(meta)) + "</span></a>";
  }

  S.views.poap = {
    mount: poMount,
    render: poRender,
    resize: function () { if (PO.ready && S.model && Math.abs((PO.main.clientWidth - 236) - (PO.tlPx || 0)) > 40) poRender(); }
  };

  // ===============================================================
  // INTEGRATION FLOWS (P2P)
  // ===============================================================
  var PP = { ready: false };

  function ppMount(container) {
    var el = h("div", { class: "poap-pp" });
    PP.kpis = h("div", { class: "grid grid-4 poap-pp-kpis" });
    el.appendChild(PP.kpis);
    var tb = h("div", { class: "poap-pp-tb" });
    PP.selTopic = h("select", { "aria-label": "Filter by topic" });
    PP.selStatus = h("select", { "aria-label": "Filter by flow status" });
    PP.selResult = h("select", { "aria-label": "Filter by result" });
    [PP.selTopic, PP.selStatus, PP.selResult].forEach(function (s) { s.addEventListener("change", ppFilterChange); });
    PP.sysBar = h("div", { class: "poap-pp-sysbar" });
    PP.reset = h("button", { type: "button", class: "btn btn-outline btn-sm", text: "Clear flow filters" });
    PP.reset.addEventListener("click", function () { S.p2pF = { topic: "", status: "", result: "", sys: [] }; ppSyncControls(); S.dirty.p2p = true; refreshAll(); });
    tb.appendChild(labelled("Topic", PP.selTopic)); tb.appendChild(labelled("Status", PP.selStatus)); tb.appendChild(labelled("Result", PP.selResult));
    tb.appendChild(PP.sysBar); tb.appendChild(PP.reset);
    el.appendChild(tb);
    var grid = h("div", { class: "poap-pp-grid" });
    PP.tableCard = h("div", { class: "card poap-pp-main" });
    PP.side = h("div", { class: "card poap-pp-side" });
    grid.appendChild(PP.tableCard); grid.appendChild(PP.side);
    el.appendChild(grid);
    container.appendChild(el);
    PP.el = el; PP.ready = true;
    PP.tableCard.addEventListener("click", function (ev) {
      var s = ev.target.closest("[data-poap-sys]");
      if (s) { ev.stopPropagation(); toggleSystem(s.getAttribute("data-poap-sys")); return; }
      var r = ev.target.closest("tr[data-flow]");
      if (r) openDrawer("flow", r.getAttribute("data-flow"));
    });
    PP.tableCard.addEventListener("keydown", function (ev) {
      if (ev.key !== "Enter" || ev.target.closest("button")) return;
      var r = ev.target.closest("tr[data-flow]"); if (r) openDrawer("flow", r.getAttribute("data-flow"));
    });
    PP.side.addEventListener("click", function (ev) {
      var s = ev.target.closest("[data-poap-sys]"); if (s) toggleSystem(s.getAttribute("data-poap-sys"));
    });
  }
  function labelled(text, ctl) { var l = h("label", { class: "poap-f" }); l.appendChild(h("span", { text: text })); l.appendChild(ctl); return l; }
  function fillSelect(sel, opts, cur, allLabel) {
    sel.innerHTML = "";
    sel.appendChild(h("option", { value: "", text: allLabel }));
    opts.forEach(function (o) { var op = h("option", { value: o, text: o }); if (o === cur) op.selected = true; sel.appendChild(op); });
    sel.value = opts.indexOf(cur) >= 0 ? cur : "";
  }
  function ppSyncControls() {
    var M = S.model; if (!M || !PP.ready) return;
    var topics = uniq(M.p2p.map(function (f) { return f.topic; }).filter(Boolean));
    var statuses = uniq(M.p2p.map(function (f) { return f.status; }).filter(Boolean));
    fillSelect(PP.selTopic, topics, S.p2pF.topic, "All topics");
    fillSelect(PP.selStatus, statuses, S.p2pF.status, "All statuses");
    fillSelect(PP.selResult, P2P_RES, S.p2pF.result, "Any result");
  }
  function ppFilterChange() {
    S.p2pF.topic = PP.selTopic.value; S.p2pF.status = PP.selStatus.value; S.p2pF.result = PP.selResult.value;
    refreshAll();
  }
  function uniq(a) { var s = {}, o = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; o.push(x); } }); return o; }
  function toggleSystem(name) {
    var i = S.p2pF.sys.indexOf(name);
    if (i >= 0) S.p2pF.sys.splice(i, 1); else S.p2pF.sys.push(name);
    S.dirty.p2p = true;
    if (S.view !== "p2p") setView("p2p"); else refreshAll();
  }
  function resTile(label, res) {
    var k = { "Pass": "green", "Fail": "red", "Blocked": "amber", "Not Run": "grey", "N/A": "dash" }[res] || "grey";
    return '<span class="poap-res poap-res-' + k + '" title="' + label + " · " + esc(res) + '"><em>' + label + "</em>" + esc(res) + "</span>";
  }

  function ppRender() {
    if (!PP.ready || !S.model) return;
    var M = S.model;
    ppSyncControls();
    var all = M.p2p;
    if (!all.length) {
      PP.kpis.innerHTML = ""; PP.sysBar.innerHTML = "";
      PP.el.querySelector(".poap-pp-tb").style.display = "none";
      PP.el.querySelector(".poap-pp-grid").style.display = "none";
      if (!PP.empty) { PP.empty = h("div", { class: "card" }); PP.el.appendChild(PP.empty); }
      PP.empty.innerHTML = emptyState("No integration flows yet", "Add point-to-point flows on the P2P_Matrix sheet (system path, test results) and reload the workbook.");
      PP.empty.style.display = "";
      S.dirty.p2p = false; return;
    }
    PP.el.querySelector(".poap-pp-tb").style.display = "";
    PP.el.querySelector(".poap-pp-grid").style.display = "";
    if (PP.empty) PP.empty.style.display = "none";
    var flows = S.fp;
    // KPIs
    var tests = 0, pass = 0, notrun = 0, bad = 0, nfail = 0, nblock = 0;
    flows.forEach(function (f) {
      [f.r2xx, f.r3xx, f.r4xx, f.r5xx].forEach(function (r) {
        if (r === "N/A") return; tests++;
        if (r === "Pass") pass++; else if (r === "Not Run") notrun++; else if (r === "Fail") { bad++; nfail++; } else if (r === "Blocked") { bad++; nblock++; }
      });
    });
    var passPct = tests ? pass / tests : 0;
    function kpi(label, val, foot, accent) {
      return '<div class="kpi-tile"><div class="accent" style="background:' + accent + '"></div><div class="kpi-label">' + label + '</div><div class="kpi-value">' + val + '</div><div class="kpi-foot">' + foot + "</div></div>";
    }
    PP.kpis.innerHTML =
      kpi("Integration flows", flows.length + (flows.length !== all.length ? '<span class="poap-kpi-of"> / ' + all.length + "</span>" : ""), plural(uniq(flows.reduce(function (a, f) { return a.concat(f.pathNodes || []); }, [])).length, "system") + " · " + plural(uniq(flows.map(function (f) { return f.topic; })).length, "topic"), "#0ea5a0") +
      kpi("Tests passed", Math.round(passPct * 100) + "%", pass + " of " + tests + " response tests" + '<div class="poap-minibar"><i style="width:' + Math.round(passPct * 100) + '%"></i></div>', "#22c55e") +
      kpi("Not run", notrun, tests ? Math.round(notrun / tests * 100) + "% of tests still to execute" : "—", "#98a2b3") +
      kpi("Failed / blocked", bad, nfail + " failed · " + nblock + " blocked", bad ? "#ef4444" : "#22c55e");
    // active system chips
    PP.sysBar.innerHTML = S.p2pF.sys.length ? '<span class="poap-sysbar-l">Touching</span>' + S.p2pF.sys.map(function (s) { return '<button type="button" class="poap-sys on" data-poap-sys="' + esc(s) + '" title="Remove ' + esc(s) + ' filter">' + esc(s) + ' <b>×</b></button>'; }).join('<span class="poap-and">and</span>') : '<span class="poap-sysbar-hint">Click a system chip to filter flows touching it</span>';
    PP.sysBar.onclick = function (ev) { var s = ev.target.closest("[data-poap-sys]"); if (s) toggleSystem(s.getAttribute("data-poap-sys")); };
    // table
    var t = '<div class="table-wrap poap-pp-wrap"><table class="data-table poap-ptable"><thead><tr><th>Flow</th><th>Topic / sub process</th><th>System path</th><th>Results</th><th>Planned</th><th>Status</th></tr></thead><tbody>';
    if (!flows.length) t += '<tr><td colspan="6" class="poap-pp-none">No flows match the current filters.</td></tr>';
    flows.forEach(function (f) {
      t += '<tr data-flow="' + esc(f.id) + '" tabindex="0" data-sel="flow:' + esc(f.id) + '"><td class="poap-fid">' + esc(f.id) + '<div class="poap-fid-s">' + esc(f.testNo) + "</div></td>" +
        '<td class="poap-ftopic"><b>' + esc(f.topic) + "</b><div>" + esc(f.subProcess) + "</div></td>" +
        '<td class="poap-fpath"><div class="poap-pathrow">' + pathChips(f.pathNodes || [], S.p2pF.sys) + "</div>" + (f.doc ? '<div class="poap-fdoc">' + esc(f.doc) + "</div>" : "") + "</td>" +
        '<td><div class="poap-resgrid">' + resTile("2xx", f.r2xx) + resTile("3xx", f.r3xx) + resTile("4xx", f.r4xx) + resTile("5xx", f.r5xx) + "</div></td>" +
        '<td class="poap-fdate">' + (f.dn != null ? fmtD(f.dn) : '<span class="poap-muted">TBC</span>') + "</td><td>" + chipHtml(f.status, statusKind(f.status)) + "</td></tr>";
    });
    t += "</tbody></table></div>";
    PP.tableCard.innerHTML = t;
    // system frequency: flows passing every filter except system
    var base = M.p2p.filter(function (f) {
      var keep = S.p2pF.sys; S.p2pF.sys = []; var ok = passFlow(f); S.p2pF.sys = keep; return ok;
    });
    var freq = {};
    base.forEach(function (f) { uniq(f.pathNodes || []).forEach(function (n) { freq[n] = (freq[n] || 0) + 1; }); });
    var names = Object.keys(freq).sort(function (a, b) { return freq[b] - freq[a] || a.localeCompare(b); });
    var mx = names.length ? freq[names[0]] : 1;
    var sh = '<div class="poap-sidehead"><div><div class="card-title">Systems by flow count</div><div class="card-sub">Click to filter · select several to find shared flows</div></div><span class="poap-count">' + names.length + "</span></div><div class=\"poap-freq\">";
    names.forEach(function (n) {
      sh += '<button type="button" class="poap-freq-i' + (S.p2pF.sys.indexOf(n) >= 0 ? " on" : "") + '" data-poap-sys="' + esc(n) + '" title="' + esc(n) + " appears in " + plural(freq[n], "flow") + '"><span class="poap-freq-n">' + esc(n) + '</span><span class="poap-freq-b"><i style="width:' + Math.round(freq[n] / mx * 100) + '%"></i></span><b>' + freq[n] + "</b></button>";
    });
    if (!names.length) sh += '<div class="poap-sideempty">No systems to show.</div>';
    sh += "</div>";
    PP.side.innerHTML = sh;
    markSelected();
    S.dirty.p2p = false;
  }

  S.views.p2p = { mount: ppMount, render: ppRender, resize: function () { /* fluid */ } };

  // ===============================================================
  // Chrome (top strip, filters), view switching, public API
  // ===============================================================
  var VIEWS = [{ key: "roadmap", label: "Roadmap" }, { key: "poap", label: "Plan on a Page" }, { key: "p2p", label: "Integration Flows" }];

  function mount(root) {
    if (!root) return;
    S.root = root;
    root.innerHTML = "";
    root.classList.add("poap-root");
    // top strip
    var top = h("div", { class: "poap-top card" });
    S.els.ribbon = h("div", { class: "poap-ribbon", style: "display:none", text: "SAMPLE DATA" });
    S.els.prog = h("h2", { class: "poap-prog", text: "Plan on a Page" });
    S.els.asat = h("div", { class: "poap-asat" });
    var ttl = h("div", { class: "poap-ttl" }, [S.els.prog, S.els.asat]);
    S.els.kpis = h("div", { class: "poap-kpis" });
    S.els.vsw = segmented(VIEWS, S.view, function (k) { setView(k); }, "Choose view");
    S.els.vsw.classList.add("poap-vsw");
    top.appendChild(S.els.ribbon);
    top.appendChild(h("div", { class: "poap-top-row" }, [ttl, S.els.vsw]));
    top.appendChild(S.els.kpis);
    root.appendChild(top);
    bindKpiClicks();
    // filters
    var fl = h("div", { class: "poap-filters card" });
    S.els.fPillar = h("select", { "aria-label": "Filter by pillar" });
    S.els.fType = h("select", { "aria-label": "Filter by activity type" });
    S.els.fStatus = h("select", { "aria-label": "Filter by status" });
    S.els.fRag = h("select", { "aria-label": "Filter by RAG" });
    [S.els.fPillar, S.els.fType, S.els.fStatus, S.els.fRag].forEach(function (s) {
      s.addEventListener("change", function () {
        setFilter({ pillar: S.els.fPillar.value, type: S.els.fType.value, status: S.els.fStatus.value, rag: S.els.fRag.value });
      });
    });
    var sb = h("div", { class: "search-box poap-search" });
    sb.innerHTML = ICON.search;
    S.els.fQ = h("input", { type: "search", placeholder: "Search activities, topics, owners…", "aria-label": "Search", autocomplete: "off" });
    var qt = null;
    S.els.fQ.addEventListener("input", function () { clearTimeout(qt); qt = setTimeout(function () { setFilter({ q: S.els.fQ.value.trim() }); }, 160); });
    sb.appendChild(S.els.fQ);
    S.els.fReset = h("button", { type: "button", class: "btn btn-outline btn-sm", text: "Reset", title: "Clear all filters" });
    S.els.fReset.addEventListener("click", function () { resetFilters(); });
    S.els.fInfo = h("div", { class: "poap-finfo" });
    fl.appendChild(labelled("Pillar", S.els.fPillar)); fl.appendChild(labelled("Activity type", S.els.fType));
    fl.appendChild(labelled("Status", S.els.fStatus)); fl.appendChild(labelled("RAG", S.els.fRag));
    fl.appendChild(sb); fl.appendChild(S.els.fReset); fl.appendChild(S.els.fInfo);
    root.appendChild(fl);
    S.els.filters = fl;
    // views
    var vw = h("div", { class: "poap-views" });
    S.els.empty = h("div", { class: "card", style: "display:none" });
    vw.appendChild(S.els.empty);
    S.els.v = {};
    VIEWS.forEach(function (v) {
      var c = h("section", { class: "poap-view", "data-view": v.key, "aria-label": v.label });
      if (v.key !== S.view) c.hidden = true;
      vw.appendChild(c); S.els.v[v.key] = c;
      S.views[v.key].mount(c);
    });
    root.appendChild(vw);
    ensureTip(); ensureDrawer();
    if (typeof ResizeObserver !== "undefined") {
      var ro = new ResizeObserver(rafThrottle(onResize));
      ro.observe(root);
      S.ro = ro;
    } else window.addEventListener("resize", rafThrottle(onResize));
    document.addEventListener("scroll", hideTip, true);
    if (S.data) render(S.data, S.ctx);
  }
  function visible() { return S.root && S.root.offsetWidth > 0 && S.root.offsetHeight > 0; }
  function onResize() {
    if (!S.model || !visible()) return;
    if (S.dirty[S.view]) { S.views[S.view].render(); return; }
    S.views[S.view].resize();
  }

  function updateChrome() {
    var M = S.model;
    S.els.ribbon.style.display = M.sample ? "" : "none";
    cls(S.root, "is-sample", M.sample);
    S.els.prog.textContent = M.program;
    S.els.asat.innerHTML = '<span class="poap-asat-d"></span>As at <b>' + fmtD(M.statusDn) + "</b>" + (M.statusDn !== M.todayDn ? ' <span class="poap-muted">· today ' + fmtD(M.todayDn, true) + "</span>" : "");
    // filter options
    var present = {}, pp = {};
    M.bars.forEach(function (b) { present[b.type] = 1; pp[b.pillar] = 1; });
    M.ms.forEach(function (m) { if (m.pillar) pp[m.pillar] = 1; });
    fillSelect(S.els.fPillar, M.pillars.map(function (p) { return p.name; }).filter(function (n) { return pp[n]; }), S.filters.pillar, "All pillars");
    fillSelect(S.els.fType, M.types.map(function (p) { return p.name; }).filter(function (n) { return present[n]; }), S.filters.type, "All types");
    var st = {}; M.bars.forEach(function (b) { st[b.status] = 1; });
    fillSelect(S.els.fStatus, STATUSES.filter(function (s) { return st[s]; }).concat(Object.keys(st).filter(function (s) { return STATUSES.indexOf(s) < 0; })), S.filters.status, "All statuses");
    fillSelect(S.els.fRag, RAGS, S.filters.rag, "All RAG");
    S.filters.pillar = S.els.fPillar.value; S.filters.type = S.els.fType.value; S.filters.status = S.els.fStatus.value; S.filters.rag = S.els.fRag.value;
    if (S.els.fQ.value.trim() !== S.filters.q) S.els.fQ.value = S.filters.q;
  }
  function updateKpis() {
    var M = S.model, k = computeKpis();
    function chip(cls2, num, label, filterKey, filterVal, title) {
      var on = filterKey && S.filters[filterKey] === filterVal;
      return '<button type="button" class="poap-kpi ' + cls2 + (on ? " on" : "") + '"' + (filterKey ? ' data-fk="' + filterKey + '" data-fv="' + esc(filterVal) + '"' : " tabindex=\"-1\"") + ' title="' + esc(title || "") + '"><b>' + num + "</b><span>" + label + "</span></button>";
    }
    var html = chip("k-bars", k.bars + (k.bars !== k.total ? '<small> / ' + k.total + "</small>" : ""), "Activities", null, null, "Bars on the roadmap (after filters)") +
      chip("k-prog", k.prog, "In progress", "status", "In Progress", "Click to filter: In Progress") +
      chip("k-done", k.done, "Complete", "status", "Complete", "Click to filter: Complete") +
      chip("k-late", k.delayed, "Delayed / slipping", "status", "Delayed", "Delayed or end date later than baseline. Click to filter: Delayed") +
      chip("k-tbc", k.tbc, "Dates TBC", "status", "Dates TBC", "Click to filter: Dates TBC");
    if (k.next) {
      var d = k.next.dn - M.statusDn;
      html += '<button type="button" class="poap-kpi k-next" data-ms="' + esc(k.next.id) + '" title="' + esc(k.next.name + " · " + fmtD(k.next.dn)) + '"><i class="poap-dia" style="background:#0f1c3f"></i><span class="poap-kn"><b>' + (d === 0 ? "Today" : d + " day" + (d === 1 ? "" : "s")) + '</b><span class="poap-kn-t">Next: ' + esc(k.next.name) + "</span></span></button>";
    } else html += '<span class="poap-kpi k-next is-none"><span class="poap-kn"><b>—</b><span class="poap-kn-t">No upcoming milestone</span></span></span>';
    S.els.kpis.innerHTML = html;
    S.els.fInfo.innerHTML = (S.view === "p2p") ? "<b>" + S.fp.length + "</b> of " + M.p2p.length + " flows" : "<b>" + k.bars + "</b> of " + k.total + " activities";
    var anyF = S.filters.pillar || S.filters.type || S.filters.status || S.filters.rag || S.filters.q || S.p2pF.topic || S.p2pF.status || S.p2pF.result || S.p2pF.sys.length;
    S.els.fReset.disabled = !anyF;
    S.els.filters.classList.toggle("is-p2p", S.view === "p2p");
  }
  function bindKpiClicks() {
    S.els.kpis.addEventListener("click", function (ev) {
      var b = ev.target.closest(".poap-kpi"); if (!b) return;
      if (b.getAttribute("data-fk")) {
        var k = b.getAttribute("data-fk"), v = b.getAttribute("data-fv"), o = {};
        o[k] = S.filters[k] === v ? "" : v; setFilter(o); return;
      }
      var ms = b.getAttribute("data-ms"); if (ms) openDrawer("ms", ms);
    });
  }

  function showEmptyAll(title, text) {
    S.model = null;
    S.els.empty.style.display = "";
    S.els.empty.innerHTML = emptyState(title, text);
    VIEWS.forEach(function (v) { S.els.v[v.key].hidden = true; });
    S.els.kpis.innerHTML = ""; S.els.fInfo.innerHTML = "";
    S.els.ribbon.style.display = "none";
    S.els.prog.textContent = "Plan on a Page"; S.els.asat.innerHTML = "";
    S.els.filters.style.display = "none"; S.els.vsw.style.display = "none";
  }

  function render(data, ctx) {
    S.data = data; S.ctx = ctx || {};
    if (!S.root) return;
    try {
      if (!data || !data.plan) { showEmptyAll("No POAP workbook loaded", "Load a workbook containing a POAP_Plan sheet (use the POAP template) to see the roadmap."); return; }
      var newData = (S.lastData !== data);
      S.lastData = data;
      S.model = buildModel(data, ctx);
      S.els.empty.style.display = "none"; S.els.filters.style.display = ""; S.els.vsw.style.display = "";
      if (newData) { S.collapsed = {}; closeDrawerSilent(); S.p2pF = { topic: "", status: "", result: "", sys: [] }; RM.inited = false; }
      updateChrome();
      applyFilters();
      updateKpis();
      S.dirty = { roadmap: true, poap: true, p2p: true };
      showView();
      if (visible()) S.views[S.view].render();
    } catch (err) {
      ATSpoap.lastError = err;
      if (window.console && console.error) console.error("[POAP] render failed", err);
      S.els.empty.style.display = ""; S.els.empty.innerHTML = emptyState("Could not render the POAP", String(err && err.message || err));
    }
  }
  function closeDrawerSilent() {
    if (drawerEl) { drawerEl.classList.remove("open"); document.body.classList.remove("poap-drawer-open"); }
    S.selKind = null; S.selId = null; S.hoverId = null;
  }
  function showView() {
    VIEWS.forEach(function (v) { S.els.v[v.key].hidden = (v.key !== S.view); });
    S.els.vsw.setActive(S.view);
    S.root.setAttribute("data-view", S.view);
  }
  function setView(v) {
    if (!S.views[v]) return;
    S.view = v;
    if (!S.model) return;
    showView();
    updateKpis();
    if (visible() && (S.dirty[v] || true)) S.views[v].render();
    hideTip();
  }
  function refreshAll() {
    if (!S.model) return;
    applyFilters(); updateKpis();
    S.dirty = { roadmap: true, poap: true, p2p: true };
    if (visible()) S.views[S.view].render();
  }
  function setFilter(part) {
    if (!part) return;
    var f = S.filters;
    ["pillar", "type", "status", "rag", "q"].forEach(function (k) { if (part[k] != null) f[k] = String(part[k]); });
    if (S.model) {
      S.els.fPillar.value = f.pillar; S.els.fType.value = f.type; S.els.fStatus.value = f.status; S.els.fRag.value = f.rag;
      if (S.els.fQ.value.trim() !== f.q) S.els.fQ.value = f.q;
    }
    if (part.topic != null) S.p2pF.topic = part.topic;
    if (part.p2pStatus != null) S.p2pF.status = part.p2pStatus;
    if (part.result != null) S.p2pF.result = part.result;
    if (part.systems) S.p2pF.sys = part.systems.slice();
    refreshAll();
  }
  function resetFilters(silentRender) {
    S.filters = { pillar: "", type: "", status: "", rag: "", q: "" };
    S.p2pF = { topic: "", status: "", result: "", sys: [] };
    if (S.model) {
      S.els.fPillar.value = ""; S.els.fType.value = ""; S.els.fStatus.value = ""; S.els.fRag.value = ""; S.els.fQ.value = "";
      applyFilters(); updateKpis();
      S.dirty = { roadmap: true, poap: true, p2p: true };
      if (!silentRender && visible()) S.views[S.view].render();
    }
  }

  // ---------------------------------------------------------------
  // Public API
  // ---------------------------------------------------------------
  // data checks for the Data Quality page: problems in POAP_Plan / POAP_Milestones that otherwise draw silently wrong bars
  ATSpoap.quality = function (data) {
    var w = [];
    if (!data) return w;
    var plan = data.plan || [], ids = {}, dupes = {}, pillars = {};
    ((data.config && data.config.pillars) || []).forEach(function (p) { pillars[p.name] = 1; });
    plan.forEach(function (b) { ids[b.id] = 1; if (b.dupOf) dupes[b.dupOf] = (dupes[b.dupOf] || 1) + 1; });
    Object.keys(dupes).forEach(function (id) { w.push({ sheet: "POAP_Plan", msg: "ID " + id + " is used on " + dupes[id] + " rows — the later rows were renamed " + id + "~2 … so dependencies point at the first one only. Give every bar its own ID." }); });
    plan.forEach(function (b) { if (b.noId) w.push({ sheet: "POAP_Plan", msg: "Row " + b.row + " (" + String(b.item || "").slice(0, 40) + ") has no ID, so it cannot be used as a dependency." }); });
    plan.forEach(function (b) {
      var nm = b.id + " (" + String(b.item || "").slice(0, 40) + ")";
      if (b.start && b.end && b.end < b.start) w.push({ sheet: "POAP_Plan", msg: nm + ": End date is before Start date." });
      if ((!b.start || !b.end) && !/tbc|tbd/i.test(String(b.status || "") + " " + String(b.start || "") + " " + String(b.end || ""))) w.push({ sheet: "POAP_Plan", msg: nm + ": Start or End date is missing, so the bar is shown as Dates TBC." });
      (b.dependsOn || []).forEach(function (d) { if (d && !ids[d]) w.push({ sheet: "POAP_Plan", msg: nm + ": depends on " + d + ", which is not in the plan." }); else if (d === b.id) w.push({ sheet: "POAP_Plan", msg: nm + ": depends on itself." }); });
      if (b.pillar && Object.keys(pillars).length && !pillars[b.pillar]) w.push({ sheet: "POAP_Plan", msg: nm + ": pillar '" + b.pillar + "' is not listed in POAP_Config > PILLARS (it gets a default colour)." });
      if (!b.topic) w.push({ sheet: "POAP_Plan", msg: nm + ": Topic is empty." });
    });
    (data.milestones || []).forEach(function (m) { if (!m.date) w.push({ sheet: "POAP_Milestones", msg: "Milestone '" + String(m.name || "").slice(0, 40) + "' has no date and is not drawn." }); });
    return w;
  };
  ATSpoap.detect = detect;
  ATSpoap.parse = parse;
  ATSpoap.mount = function (root) { mount(root); };
  ATSpoap.render = render;
  ATSpoap.setFilter = setFilter;
  ATSpoap.getFilters = function () { return JSON.parse(JSON.stringify({ filters: S.filters, p2p: S.p2pF })); };
  ATSpoap.resetFilters = function () { resetFilters(); };
  ATSpoap.setView = setView;
  ATSpoap.getView = function () { return S.view; };
  ATSpoap.focus = function (id) { return S.model ? (S.model.msById[id] ? focusMs(id) : focusBar(id)) : false; };
  ATSpoap.openDetails = function (kind, id) { openDrawer(kind, id); };
  ATSpoap.closeDetails = closeDrawer;
  ATSpoap.setZoom = function (k) { if (RM.ready && S.model) { if (k === "fit") rmFit(); else rmSetZoom(k); } };
  ATSpoap._state = S; // for diagnostics / tests
})();
