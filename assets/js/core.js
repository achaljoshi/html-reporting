/* ============================================================
   ATS Program Dashboard — core library (namespace window.ATS)
   Dates · Excel helpers · charts · data store · folder loader
   100% offline. No network calls.
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});

  // ------------------------------------------------------------------
  // DOM helpers
  // ------------------------------------------------------------------
  const $ = (s, r) => (r || document).querySelector(s);
  const $all = (s, r) => Array.from((r || document).querySelectorAll(s));
  const esc = (s) =>
    String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  ATS.dom = { $, $all, esc };

  // ------------------------------------------------------------------
  // Dates — everything is an ISO "YYYY-MM-DD" string; math is done at local midnight
  // ------------------------------------------------------------------
  const pad = (n) => String(n).padStart(2, "0");
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const MONTHS_LONG = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

  const D = {
    pad,
    iso(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); },
    parse(iso) {
      const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso || "");
      return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
    },
    todayIso() { return D.iso(new Date()); },
    // any cell value -> ISO or null
    toIso(v) {
      if (v == null || v === "") return null;
      if (v instanceof Date) {
        if (isNaN(v)) return null;
        return D.iso(new Date(v.getTime() + 12 * 3600 * 1000)); // tolerate +/-12h timezone skew
      }
      if (typeof v === "number") {
        if (v > 20000 && v < 90000) {
          const d = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
          return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate());
        }
        return null;
      }
      let s = String(v).trim();
      if (!s) return null;
      if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
      s = s.replace(/^[A-Za-z]{3,9},?\s+/, ""); // "Tue 16/06/26"
      let m = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/.exec(s);
      if (m) {
        let y = +m[3];
        if (y < 100) y += 2000;
        return y + "-" + pad(+m[2]) + "-" + pad(+m[1]);
      }
      const t = Date.parse(s);
      return isNaN(t) ? null : D.iso(new Date(t));
    },
    add(iso, n) { const d = D.parse(iso); d.setDate(d.getDate() + n); return D.iso(d); },
    diff(a, b) { return Math.round((D.parse(b) - D.parse(a)) / 86400000); }, // b - a in days
    dow(iso) { return D.parse(iso).getDay(); },
    monthKey(iso) { return iso.slice(0, 7); },
    monthStart(key) { return key + "-01"; },
    monthEnd(key) { const [y, m] = key.split("-").map(Number); return D.iso(new Date(y, m, 0)); },
    prevMonthKey(key) { const [y, m] = key.split("-").map(Number); const d = new Date(y, m - 2, 1); return d.getFullYear() + "-" + pad(d.getMonth() + 1); },
    nextMonthKey(key) { const [y, m] = key.split("-").map(Number); const d = new Date(y, m, 1); return d.getFullYear() + "-" + pad(d.getMonth() + 1); },
    quarterOf(iso) {
      const y = +iso.slice(0, 4), m = +iso.slice(5, 7), q = Math.floor((m - 1) / 3);
      return { start: y + "-" + pad(q * 3 + 1) + "-01", end: D.monthEnd(y + "-" + pad(q * 3 + 3)), label: "Q" + (q + 1) + " " + y, q: q + 1, y };
    },
    // Friday on/after date
    weekEnding(iso) { const dw = D.dow(iso); return D.add(iso, (5 - dw + 7) % 7); },
    fmt(iso, withYear = true) {
      const d = D.parse(iso); if (!d) return "—";
      return d.getDate() + " " + MONTHS[d.getMonth()] + (withYear ? " " + d.getFullYear() : "");
    },
    fmtMonth(key, long = true) {
      if (!key) return "—";
      const [y, m] = key.split("-").map(Number);
      return (long ? MONTHS_LONG : MONTHS)[m - 1] + " " + y;
    },
    // Excel NETWORKDAYS (inclusive both ends, Mon–Fri, minus holidays)
    networkdays(a, b, hols) {
      if (!a || !b) return 0;
      let sign = 1;
      if (a > b) { const t = a; a = b; b = t; sign = -1; }
      let n = 0;
      const d = D.parse(a), end = D.parse(b);
      while (d <= end) {
        const dw = d.getDay();
        if (dw !== 0 && dw !== 6 && !(hols && hols.has(D.iso(d)))) n++;
        d.setDate(d.getDate() + 1);
      }
      return n * sign;
    },
    maxIso(a, b) { return !a ? b : !b ? a : a > b ? a : b; },
    minIso(a, b) { return !a ? b : !b ? a : a < b ? a : b; },
  };
  ATS.date = D;

  // ------------------------------------------------------------------
  // Number / text helpers
  // ------------------------------------------------------------------
  const N = {
    num(v) {
      if (v == null || v === "") return null;
      if (typeof v === "number") return isNaN(v) ? null : v;
      let s = String(v).trim().replace(/[£,\s]/g, "");
      if (!s) return null;
      const pct = s.endsWith("%");
      if (pct) s = s.slice(0, -1);
      const n = parseFloat(s);
      if (isNaN(n)) return null;
      return pct ? n / 100 : n;
    },
    ratio(v, dflt) { const n = N.num(v); if (n == null) return dflt; return n > 1 ? n / 100 : n; },
    pct(v, dp = 0) { return v == null || isNaN(v) ? "—" : (v * 100).toFixed(dp) + "%"; },
    pct1(v) { return N.pct(v, 1); },
    gbp(v) { return v == null ? "—" : "£" + Math.round(v).toLocaleString("en-GB"); },
    gbpK(v) { return v == null ? "—" : v >= 1e6 ? "£" + (v / 1e6).toFixed(2) + "m" : "£" + Math.round(v / 1000) + "k"; },
    int(v) { return v == null ? "—" : Math.round(v).toLocaleString("en-GB"); },
    round1(v) { return Math.round(v * 10) / 10; },
    str(v) { return v == null ? "" : String(v).trim(); },
    norm(s) { return String(s == null ? "" : s).toLowerCase().replace(/[^a-z0-9]+/g, ""); },
    plural(n, one, many) { return n + " " + (n === 1 ? one : many || one + "s"); },
  };
  ATS.num = N;

  // RAG
  const RAGS = { Green: "#22c55e", Amber: "#f5a524", Red: "#ef4444", Grey: "#98a2b3" };
  const RAG_ORDER = { Grey: 0, Green: 1, Amber: 2, Red: 3 };
  const R = {
    color: (r) => RAGS[r] || RAGS.Grey,
    worst(list) { let w = "Grey"; (list || []).forEach((r) => { if (r && RAG_ORDER[r] > RAG_ORDER[w]) w = r; }); return w; },
    norm(v) {
      const s = N.str(v).toLowerCase();
      if (!s) return "";
      if (s.startsWith("g")) return "Green";
      if (s.startsWith("a")) return "Amber";
      if (s.startsWith("r")) return "Red";
      return "";
    },
    cls: (r) => ({ Green: "green", Amber: "amber", Red: "red" }[r] || "grey"),
    chip(r, text) { return `<span class="chip ${R.cls(r)}">${esc(text || (r === "Grey" ? "N/A" : r))}</span>`; },
    dot(r) { return `<span class="sx-dot" style="background:${R.color(r)}"></span>`; },
  };
  ATS.rag = R;

  // ------------------------------------------------------------------
  // Excel helpers (SheetJS)
  // ------------------------------------------------------------------
  const X = {
    aoa(ws) {
      return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: true });
    },
    // find header row (best match of wanted normalized names) in first N rows
    findHeader(aoa, wanted, maxScan = 15) {
      let best = null;
      const want = wanted.map(N.norm);
      for (let r = 0; r < Math.min(maxScan, aoa.length); r++) {
        const row = aoa[r] || [];
        const map = {};
        let hits = 0;
        row.forEach((cell, c) => {
          const k = N.norm(cell);
          if (!k) return;
          if (!(k in map)) map[k] = c;
          if (want.includes(k)) hits++;
        });
        if (hits >= 2 && (!best || hits > best.hits)) best = { row: r, map, hits };
      }
      return best;
    },
    // -> array of row objects keyed by normalized header; skips blank + EXAMPLE rows
    table(wb, sheetName, wanted) {
      const ws = wb.Sheets[sheetName];
      if (!ws) return [];
      const aoa = X.aoa(ws);
      const hdr = X.findHeader(aoa, wanted);
      if (!hdr) return [];
      const keys = Object.keys(hdr.map);
      const out = [];
      for (let r = hdr.row + 1; r < aoa.length; r++) {
        const row = aoa[r] || [];
        let any = false;
        const obj = { _row: r + 1 };
        keys.forEach((k) => { const v = row[hdr.map[k]]; obj[k] = v === undefined ? null : v; if (v != null && v !== "") any = true; });
        if (!any) continue;
        const notes = N.str(obj.notes).toLowerCase();
        if (notes.startsWith("example")) continue;
        out.push(obj);
      }
      return out;
    },
    // first value whose normalized key starts with any prefix
    field(row, ...prefixes) {
      for (const p of prefixes) {
        const pn = N.norm(p);
        if (pn in row && row[pn] != null && row[pn] !== "") return row[pn];
      }
      for (const p of prefixes) {
        const pn = N.norm(p);
        for (const k in row) if (k.startsWith(pn) && row[k] != null && row[k] !== "") return row[k];
      }
      return null;
    },
  };
  ATS.xl = X;

  // ------------------------------------------------------------------
  // Charts (Chart.js) — keyed registry so re-renders never leak
  // ------------------------------------------------------------------
  const charts = {};
  ATS.chart = function (key, canvas, config) {
    if (charts[key]) { try { charts[key].destroy(); } catch (e) {} }
    if (!canvas) return null;
    charts[key] = new Chart(canvas, config);
    return charts[key];
  };
  ATS.destroyCharts = function (prefix) {
    Object.keys(charts).forEach((k) => {
      if (!prefix || k.startsWith(prefix)) { try { charts[k].destroy(); } catch (e) {} delete charts[k]; }
    });
  };
  ATS.chartBase = function () {
    return {
      maintainAspectRatio: false,
      animation: { duration: 350 },
      plugins: { legend: { labels: { boxWidth: 10, boxHeight: 10, usePointStyle: true, font: { size: 11 }, color: "#475467" } }, tooltip: { backgroundColor: "#0f1c3f", padding: 10, cornerRadius: 8 } },
    };
  };

  // ------------------------------------------------------------------
  // Events + data store (many portfolios; store.kpi/poap/raid = the SELECTED portfolio)
  // ------------------------------------------------------------------
  const listeners = {};
  ATS.on = (ev, fn) => { (listeners[ev] = listeners[ev] || []).push(fn); };
  ATS.emit = (ev, payload) => { (listeners[ev] || []).forEach((fn) => { try { fn(payload); } catch (e) { console.error(e); } }); };

  ATS.VER = 5; // bump when parsed-data shape changes (invalidates cached data)
  ATS.store = { portfolios: {}, order: [], current: null, kpi: null, poap: null, raid: null, meta: { source: "none", loadedAt: null, files: [] } };

  ATS.portfolioNames = () => ATS.store.order.slice();
  ATS.hasData = () => ATS.store.order.length > 0;
  ATS.setPortfolio = function (name, silent) {
    const st = ATS.store, p = st.portfolios[name];
    if (!p) return false;
    st.current = name; st.kpi = p.kpi || null; st.poap = p.poap || null; st.raid = p.raid || null;
    if (!silent) { persist(); ATS.emit("portfolio", name); }
    return true;
  };
  function applyPortfolios(list, meta) {
    const st = ATS.store, prev = st.current;
    st.portfolios = {}; st.order = [];
    list.slice().sort((a, b) => a.name.localeCompare(b.name)).forEach((p) => { st.portfolios[p.name] = { kpi: p.kpi || null, poap: p.poap || null, raid: p.raid || null, dir: p.dir || "" }; st.order.push(p.name); });
    st.meta = meta;
    const pick = prev && st.portfolios[prev] ? prev : st.order[0] || null;
    if (pick) ATS.setPortfolio(pick, true); else { st.current = null; st.kpi = st.poap = st.raid = null; }
  }

  // ---- persistence: IndexedDB (localStorage is too small for 16+ portfolios)
  const IDB = "ats-handles";
  function idb() {
    return new Promise((res, rej) => {
      const r = indexedDB.open(IDB, 2);
      r.onupgradeneeded = () => { const db = r.result; if (!db.objectStoreNames.contains("h")) db.createObjectStore("h"); if (!db.objectStoreNames.contains("data")) db.createObjectStore("data"); };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  function persist() {
    const st = ATS.store;
    const payload = { ver: ATS.VER, portfolios: st.portfolios, order: st.order, current: st.current, meta: st.meta };
    idb().then((db) => { db.transaction("data", "readwrite").objectStore("data").put(payload, "program"); }).catch((e) => console.warn("Could not persist dataset", e));
  }
  ATS.loadPersisted = async function () {
    try {
      const db = await idb();
      const p = await new Promise((res) => { const q = db.transaction("data").objectStore("data").get("program"); q.onsuccess = () => res(q.result || null); q.onerror = () => res(null); });
      if (!p || p.ver !== ATS.VER || !p.order || !p.order.length) return false;
      const st = ATS.store;
      st.portfolios = p.portfolios; st.order = p.order; st.meta = p.meta || st.meta;
      ATS.setPortfolio(p.current && st.portfolios[p.current] ? p.current : st.order[0], true);
      return true;
    } catch (e) { return false; }
  };
  ATS.clearPersisted = async function () { try { const db = await idb(); db.transaction("data", "readwrite").objectStore("data").delete("program"); } catch (e) {} };

  ATS.useSample = function () {
    const s = window.ATS_SAMPLE;
    if (!s || !s.portfolios || !s.portfolios.length) return false;
    applyPortfolios(s.portfolios, { source: "sample", loadedAt: new Date().toISOString(), files: s.files || [] });
    persist();
    ATS.emit("data");
    return true;
  };
  // Display marker for demo data is switched off on purpose (no "SAMPLE DATA" labels anywhere); data origin is still in store.meta.source
  ATS.isSample = function () { return false; };

  // ------------------------------------------------------------------
  // Grouping parsed workbooks into portfolios (pure — also used by tools/build_sample_js.js)
  //   item = { kind:'kpi'|'poap'|'raid', parsed, dir, fileName, ts, sz }
  // A portfolio = one weekly workbook. Its POAP/RAID are the ones in the same folder
  // (matched by portfolio name when a folder holds several weekly workbooks).
  // ------------------------------------------------------------------
  const nrm = N.norm;
  const better = (a, b) => ((a.sz > 0) !== (b.sz > 0) ? a.sz > 0 : a.ts >= b.ts);
  function best(list) { return list.reduce((m, x) => (!m || better(x, m) ? x : m), null); }
  function intName(it) { return it.kind === "poap" ? (it.parsed.config && it.parsed.config.program) || "" : it.kind === "raid" ? (it.parsed.info && it.parsed.info.project) || "" : ""; }
  function matchesName(it, name) {
    const n = nrm(name); if (!n) return false;
    return nrm(it.fileName).includes(n) || nrm(intName(it)).includes(n);
  }
  ATS.portfolioName = function (kpi, fileName, dir) {
    const c = kpi.config || {}, ok = (x) => x && !/^</.test(x);
    if (ok(c.portfolio)) return c.portfolio;                                   // 1. Config > Portfolio Name
    const folder = dir && dir.split("/").pop();
    if (folder) return folder;                                                 // 2. its folder name
    const fm = /ATS_Weekly_Data\s*[-–]\s*(.+?)\.xlsx?$/i.exec(fileName || "");
    if (fm) return fm[1].trim();                                               // 3. "ATS_Weekly_Data - <Portfolio>.xlsx"
    if (/template/i.test(fileName || "")) return "(blank template)";
    const m = /[–—-]\s*([^–—-]+)$/.exec(c.programName || "");                    // 4. tail of Program Name
    return m ? m[1].trim() : c.programName || String(fileName || "Portfolio").replace(/\.xlsx?$/i, "");
  };
  ATS.groupPortfolios = function (items) {
    // blank templates (no rows) are ignored whenever a workbook of the same kind has real data
    ["kpi", "poap", "raid"].forEach((k) => { if (items.some((i) => i.kind === k && i.sz > 0)) items = items.filter((i) => i.kind !== k || i.sz > 0); });
    const byDir = {};
    items.forEach((it) => { (byDir[it.dir || ""] = byDir[it.dir || ""] || []).push(it); });
    const out = [], used = new Set(), orphans = { poap: [], raid: [] };
    Object.keys(byDir).sort().forEach((dir) => {
      const list = byDir[dir], kpis = list.filter((i) => i.kind === "kpi");
      const cand = { poap: list.filter((i) => i.kind === "poap"), raid: list.filter((i) => i.kind === "raid") };
      if (!kpis.length) { orphans.poap.push(...cand.poap); orphans.raid.push(...cand.raid); return; }
      const byName = {};
      kpis.forEach((k) => { const nm = ATS.portfolioName(k.parsed, k.fileName, dir); if (!byName[nm] || better(k, byName[nm])) byName[nm] = k; });
      const names = Object.keys(byName);
      names.forEach((nm) => {
        const pick = (kind) => {
          const c = cand[kind]; if (!c.length) return null;
          if (names.length === 1) return best(c);
          const m = c.filter((x) => matchesName(x, nm)); return m.length ? best(m) : null;
        };
        const poap = pick("poap"), raid = pick("raid");
        let name = nm, i = 2; while (used.has(name)) name = nm + " (" + i++ + ")"; used.add(name);
        out.push({ name, dir, kpi: byName[nm].parsed, poap: poap && poap.parsed, raid: raid && raid.parsed, _items: [byName[nm], poap, raid].filter(Boolean) });
      });
    });
    // POAP / RAID files sitting in folders without a weekly workbook: attach by name, or to the only portfolio lacking one
    ["poap", "raid"].forEach((kind) => {
      orphans[kind].forEach((o) => {
        let target = out.find((p) => !p[kind] && matchesName(o, p.name));
        if (!target) { const lacking = out.filter((p) => !p[kind]); if (lacking.length === 1 && orphans[kind].length === 1) target = lacking[0]; }
        if (target) { target[kind] = o.parsed; target._items.push(o); }
      });
    });
    return out;
  };

  // ------------------------------------------------------------------
  // Loading workbooks
  // ------------------------------------------------------------------
  function readBuffer(file) {
    if (file.arrayBuffer) return file.arrayBuffer();
    return new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = rej; fr.readAsArrayBuffer(file); });
  }
  function dirOf(f) {
    let p = f._path || f.webkitRelativePath || "";
    if (!f._path && p) p = p.split("/").slice(1).join("/"); // strip the picked root folder name
    const parts = p.split("/"); parts.pop();
    return parts.join("/");
  }

  // files: File objects from a folder picker / file input / directory handle
  ATS.loadFiles = async function (files, label, opts) {
    opts = opts || {};
    const xl = Array.from(files).filter((f) => /\.xlsx?$/i.test(f.name) && !f.name.startsWith("~$"));
    const items = [], report = [];
    for (const f of xl) {
      try {
        const buf = await readBuffer(f);
        const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
        let kind = null;
        if (ATS.kpi && ATS.kpi.detect(wb)) kind = "kpi";
        else if (ATS.poap && ATS.poap.detect && ATS.poap.detect(wb)) kind = "poap";
        else if (ATS.raid && ATS.raid.detect(wb)) kind = "raid";
        if (!kind) { report.push({ name: f.name, kind: "ignored", note: "Not a recognised ATS workbook" }); continue; }
        const mod = ATS[kind], parsed = mod.parse(wb);
        if (!parsed) { report.push({ name: f.name, kind, note: "Could not read" }); continue; }
        const sz = kind === "kpi" ? parsed.snapshots.length + parsed.defects.length : kind === "poap" ? (parsed.plan || []).length : parsed.risks.length + parsed.issues.length;
        items.push({ kind, parsed, dir: dirOf(f), fileName: f.name, ts: f.lastModified || 0, sz });
        report.push({ name: f.name, kind, dir: dirOf(f), note: mod.describe ? mod.describe(parsed) : "OK" });
      } catch (e) {
        console.error("Failed to read", f.name, e);
        report.push({ name: f.name, kind: "error", note: String(e.message || e) });
      }
    }
    // ignore empty blank templates when a real workbook of the same kind exists anywhere
    const groups = ATS.groupPortfolios(items);
    if (!groups.length) return { ok: false, report };
    let list = groups;
    if (opts.merge && ATS.store.order.length) {
      // "choose files": merge into what is already loaded (replace same-named portfolios, keep the rest)
      const keep = ATS.store.order.filter((n) => !groups.some((g) => g.name === n)).map((n) => Object.assign({ name: n }, ATS.store.portfolios[n]));
      list = keep.concat(groups);
    }
    const files2 = [];
    groups.forEach((g) => (g._items || []).forEach((it) => files2.push({ name: (g.dir ? g.dir + "/" : "") + it.fileName, kind: it.kind, portfolio: g.name, note: (ATS[it.kind].describe ? ATS[it.kind].describe(it.parsed) : "") })));
    applyPortfolios(list, { source: "folder", label: label || "", loadedAt: new Date().toISOString(), files: files2 });
    persist();
    ATS.emit("data");
    return { ok: true, report, portfolios: groups.map((g) => g.name) };
  };

  // ---- directory handles (Chrome/Edge): remembered so "Refresh" needs no re-pick ----
  async function saveHandle(h) { try { const db = await idb(); db.transaction("h", "readwrite").objectStore("h").put(h, "program"); } catch (e) {} }
  async function getHandle() {
    try { const db = await idb(); return await new Promise((res) => { const q = db.transaction("h").objectStore("h").get("program"); q.onsuccess = () => res(q.result || null); q.onerror = () => res(null); }); }
    catch (e) { return null; }
  }
  async function filesFromHandle(dir, prefix, depth) {
    prefix = prefix || ""; depth = depth || 0;
    const out = [];
    for await (const [name, h] of dir.entries()) {
      if (h.kind === "file" && /\.xlsx?$/i.test(name)) { const f = await h.getFile(); f._path = prefix + name; out.push(f); }
      else if (h.kind === "directory" && depth < 3 && !name.startsWith(".")) out.push(...(await filesFromHandle(h, prefix + name + "/", depth + 1)));
    }
    return out;
  }
  ATS.canResync = async function () { return !!(window.showDirectoryPicker && (await getHandle())); };

  ATS.pickFolder = async function () {
    if (window.showDirectoryPicker) {
      try {
        const dir = await window.showDirectoryPicker({ mode: "read" });
        const files = await filesFromHandle(dir);
        await saveHandle(dir);
        return await ATS.loadFiles(files, dir.name);
      } catch (e) {
        if (e && e.name === "AbortError") return { ok: false, cancelled: true, report: [] };
        console.warn("Directory picker unavailable, falling back", e);
      }
    }
    return new Promise((resolve) => {
      const inp = $("#ats-folder-input");
      inp.onchange = async () => {
        if (!inp.files.length) return resolve({ ok: false, cancelled: true, report: [] });
        const first = inp.files[0].webkitRelativePath || "";
        resolve(await ATS.loadFiles(inp.files, first.split("/")[0]));
        inp.value = "";
      };
      inp.click();
    });
  };
  ATS.pickFiles = function () {
    return new Promise((resolve) => {
      const inp = $("#ats-files-input");
      inp.onchange = async () => {
        if (!inp.files.length) return resolve({ ok: false, cancelled: true, report: [] });
        resolve(await ATS.loadFiles(inp.files, "Selected files", { merge: true }));
        inp.value = "";
      };
      inp.click();
    });
  };
  // Re-read the remembered folder without a picker (one permission click per browser session)
  ATS.resync = async function () {
    const h = await getHandle();
    if (!h) return ATS.pickFolder();
    try {
      let perm = await h.queryPermission({ mode: "read" });
      if (perm !== "granted") perm = await h.requestPermission({ mode: "read" });
      if (perm !== "granted") return { ok: false, report: [], denied: true };
      return await ATS.loadFiles(await filesFromHandle(h), h.name);
    } catch (e) {
      console.warn("Resync failed, falling back to picker", e);
      return ATS.pickFolder();
    }
  };

  // ------------------------------------------------------------------
  // Export helpers (page -> PNG, chart -> PNG)
  // ------------------------------------------------------------------
  ATS.toast = function (msg, kind) {
    let t = $("#ats-toast");
    if (!t) { t = document.createElement("div"); t.id = "ats-toast"; document.body.appendChild(t); }
    t.className = "sx-toast show " + (kind || "");
    t.textContent = msg;
    clearTimeout(ATS._toastT);
    ATS._toastT = setTimeout(() => t.classList.remove("show"), 2800);
  };
  ATS.elementToBlob = async function (el) {
    if (!window.html2canvas) throw new Error("Image export library missing");
    const canvas = await window.html2canvas(el, { backgroundColor: "#ffffff", scale: 2, useCORS: true, logging: false });
    return await new Promise((res) => canvas.toBlob(res, "image/png"));
  };
  ATS.copyElementImage = async function (el, filename) {
    try {
      const blob = await ATS.elementToBlob(el);
      if (navigator.clipboard && window.ClipboardItem) {
        await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
        ATS.toast("Copied as picture — paste straight into PowerPoint", "ok");
        return;
      }
      throw new Error("clipboard unavailable");
    } catch (e) {
      try { await ATS.downloadElementImage(el, filename); } catch (e2) { ATS.toast("Could not export image", "err"); }
    }
  };
  ATS.downloadElementImage = async function (el, filename) {
    const blob = await ATS.elementToBlob(el);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = (filename || "ats-export") + ".png";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    ATS.toast("Saved " + a.download, "ok");
  };
})();
