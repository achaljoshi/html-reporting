/* ============================================================
   Shared UI builders for the program sections
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date, N = ATS.num, R = ATS.rag, K = ATS.kpi, esc = ATS.dom.esc;
  const U = (ATS.ui = {});

  U.C = { teal: "#0ea5a0", navy: "#1b2a4a", indigo: "#6366f1", amber: "#f5a524", red: "#ef4444", green: "#22c55e", slate: "#94a3b8", sky: "#38bdf8", violet: "#8b5cf6", grid: "#eef1f6" };
  U.ragLetter = (r) => ({ Green: "G", Amber: "A", Red: "R" }[r] || "–");

  U.ragBadge = function (res, label) {
    const ov = res.ragOverridden ? `<span class="ov" title="Computed from data: ${esc(res.autoRag)}">overridden</span>` : "";
    return `<div class="sx-rag-badge"><div class="big" style="background:${R.color(res.rag)}">${U.ragLetter(res.rag)}</div><small>${esc(label || (res.rag === "Grey" ? "N/A" : res.rag))}</small>${ov}</div>`;
  };

  U.deltaChip = function (res, label) {
    const d = K.delta(res);
    if (!d) return "";
    const cls = d.good == null ? "neu" : d.good ? "good" : "bad";
    return `<span class="sx-delta ${cls}" title="vs ${esc(res.prev ? "previous period" : "")}">${esc(d.text)}${label ? " " + esc(label) : ""}</span>`;
  };

  U.list = (arr) => (arr && arr.length ? `<ul>${arr.map((x) => `<li>${esc(x)}</li>`).join("")}</ul>` : `<div class="sx-empty-note">Nothing to report.</div>`);

  U.narrative = function (res) {
    const t = (k) => (res.src && res.src[k] === "excel" ? `<span class="tag excel" title="Written in the Commentary sheet of your workbook">From Excel</span>` : `<span class="tag" title="Drafted automatically from the numbers — add a row in the Commentary sheet to override">Auto-drafted</span>`);
    return `
      <div class="sx-panel hi"><h5>Key Highlights ${t("highlights")}</h5>${U.list(res.highlights)}</div>
      <div class="sx-panel co"><h5>Area of Concern ${t("concern")}</h5>${U.list(res.concern)}</div>
      <div class="sx-panel ac"><h5>Actions Underway ${t("actions")}</h5>${U.list(res.actions)}</div>
      <div class="sx-panel ex"><h5>Executive Commentary ${t("exec")}</h5><p>${esc(res.exec || "—")}</p></div>`;
  };

  // generic slide wrapper
  U.slide = function (o) {
    const id = o.id || "sx-slide-" + Math.random().toString(36).slice(2, 7);
    return `
      <article class="sx-slide" id="${id}">
        <header class="sx-slide-head">
          <div><div class="sx-slide-eyebrow">${esc(o.eyebrow || "")}</div><h3>${esc(o.title)}</h3>${o.sub ? `<div class="sx-slide-sub">${o.sub}</div>` : ""}</div>
          ${o.badge || ""}
        </header>
        <div class="${o.single ? "" : "sx-body"}">${o.single ? o.left : `<div class="sx-left">${o.left}</div><div class="sx-right">${o.right || ""}</div>`}</div>
        <footer class="sx-slide-foot"><span>${o.foot || ""}</span>
          <span class="sx-foot-btns" data-export="${id}">
            <button class="btn btn-outline btn-sm" data-act="copy">Copy as picture</button>
            <button class="btn btn-outline btn-sm" data-act="png">Download PNG</button>
          </span></footer>
      </article>`;
  };
  // wire the export buttons inside root
  U.wireExports = function (root, fname) {
    root.querySelectorAll("[data-export]").forEach((bar) => {
      const el = document.getElementById(bar.getAttribute("data-export"));
      bar.querySelectorAll("button").forEach((b) => {
        b.onclick = async () => {
          bar.style.visibility = "hidden";
          try { await (b.dataset.act === "copy" ? ATS.copyElementImage(el, fname) : ATS.downloadElementImage(el, fname)); }
          finally { bar.style.visibility = ""; }
        };
      });
    });
  };

  U.stat = (value, label, extra) => `<div class="sx-stat"><b>${value}</b><span>${esc(label)}</span>${extra || ""}</div>`;
  U.chartBox = (id, h, title, sub) => `<div>${title ? `<div class="sx-ctitle">${esc(title)}${sub ? ` <span class="sx-csub">· ${esc(sub)}</span>` : ""}</div>` : ""}<div class="sx-chartbox" style="height:${h}px"><canvas id="${id}"></canvas></div></div>`;

  U.sparkline = function (vals, color, w, h) {
    w = w || 96; h = h || 28;
    const v = (vals || []).filter((x) => x != null && !isNaN(x));
    if (v.length < 2) return `<svg class="sx-spark" width="${w}" height="${h}"></svg>`;
    const mn = Math.min(...v), mx = Math.max(...v), rng = mx - mn || 1;
    const pts = v.map((x, i) => [(i / (v.length - 1)) * (w - 4) + 2, h - 3 - ((x - mn) / rng) * (h - 8)]);
    const path = pts.map((p, i) => (i ? "L" : "M") + p[0].toFixed(1) + " " + p[1].toFixed(1)).join(" ");
    const last = pts[pts.length - 1];
    return `<svg class="sx-spark" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><path d="${path}" fill="none" stroke="${color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><circle cx="${last[0]}" cy="${last[1]}" r="3" fill="${color}"/></svg>`;
  };

  U.ring = function (pct, color, size, label) {
    size = size || 96;
    const r = size / 2 - 8, c = 2 * Math.PI * r, p = Math.max(0, Math.min(1, pct || 0));
    return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}"><circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="rgba(255,255,255,.15)" stroke-width="9"/>
      <circle cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="${color}" stroke-width="9" stroke-linecap="round" stroke-dasharray="${(c * p).toFixed(1)} ${c.toFixed(1)}" transform="rotate(-90 ${size / 2} ${size / 2})"/>
      <text x="50%" y="50%" text-anchor="middle" dominant-baseline="central" fill="#fff" font-size="${size * 0.28}" font-weight="800">${esc(label)}</text></svg>`;
  };

  U.tile = function (res, withSpark) {
    const d = U.deltaChip(res);
    const val = res.metric ? res.metric.text : res.empty ? "—" : "";
    return `
      <button class="sx-tile" data-kpi="${res.key}">
        <span class="glow" style="background:${R.color(res.rag)}"></span>
        <div class="t-top"><div class="t-name">${esc(res.name)}</div>${R.chip(res.rag)}</div>
        <div class="t-val">${esc(val)}</div>
        <div class="t-pos">${esc(res.position)}</div>
        <div class="t-foot"><span class="cad">${esc(res.cadence)}</span>${withSpark ? U.sparkline(withSpark.vals, R.color(res.rag)) : d}</div>
      </button>`;
  };

  U.table = function (heads, rows, opts) {
    opts = opts || {};
    return `<div class="table-wrap"><table class="data-table"><thead><tr>${heads.map((h) => `<th style="cursor:default">${esc(h)}</th>`).join("")}</tr></thead><tbody>${rows.map((r) => `<tr>${r.map((c) => `<td class="${opts.wrap ? "sx-wrap" : ""}">${c}</td>`).join("")}</tr>`).join("") || `<tr><td colspan="${heads.length}" style="text-align:center;color:var(--sub);padding:18px">No data</td></tr>`}</tbody></table></div>`;
  };

  // common chart option fragments
  U.axes = (o) => Object.assign({ x: { grid: { display: false }, ticks: { color: "#667085", font: { size: 11 } } }, y: { grid: { color: U.C.grid }, ticks: { color: "#667085", font: { size: 11 } }, beginAtZero: true } }, o || {});
  U.base = () => ATS.chartBase();
  U.shortWeek = (we) => D.fmt(we, false);

  // simple tab-nav for pages
  U.noData = (what) => `<div class="empty-state" style="padding:50px 20px"><h2>${esc(what || "No data")}</h2><p>Nothing to show for this period yet. Add the rows in the Excel workbook and refresh.</p></div>`;
})();
