/* ============================================================
   Attachments UI — chips next to notes, the Evidence viewer (image lightbox, email reader, files) and the Evidence library page.
   Anything can render   ATS.att.chip(ids, label)   -> a clickable paperclip; clicks anywhere on [data-att] open the viewer.
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const A = (ATS.att = ATS.att || {});
  const esc = ATS.dom.esc;
  const KIND = { image: "Picture", msg: "Outlook email", file: "File" };
  const ICON = {
    clip: '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 11.5l-8.6 8.6a5 5 0 01-7.1-7.1l8.9-8.9a3.3 3.3 0 014.7 4.7l-8.9 8.9a1.7 1.7 0 01-2.4-2.4l8.2-8.2"/></svg>',
    mail: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3 7l9 6 9-6"/></svg>',
    img: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="M21 16l-5-5-8 8"/></svg>',
    file: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 00-2 2v14a2 2 0 002 2h10a2 2 0 002-2V8z"/><path d="M14 3v5h5"/></svg>',
    dl: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12M7 10l5 5 5-5M4 20h16"/></svg>',
  };
  const iconFor = (it) => (it.kind === "msg" || it.email ? ICON.mail : it.kind === "image" ? ICON.img : ICON.file);
  const kindOf = (it) => (it.email || it.kind === "msg" ? "msg" : /^image\/(png|jpe?g|gif|bmp|webp|svg\+xml)$/.test(it.mime || "") || it.kind === "image" ? "image" : "file");
  const dataUrl = (it) => "data:" + A.safeMime(it.mime) + ";base64," + it.b64;
  const bytes = (n) => (n < 1024 ? n + " B" : n < 1048576 ? (n / 1024).toFixed(n < 10240 ? 1 : 0) + " KB" : (n / 1048576).toFixed(1) + " MB");
  const when = (iso) => { if (!iso) return ""; const d = new Date(iso); return isNaN(d) ? "" : d.toLocaleString("en-GB", { weekday: "short", day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }); };
  const hue = (s) => { let h = 0; for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) % 360; return h; };
  const initials = (n) => (String(n || "?").replace(/[<>"]/g, "").split(/[\s@.]+/).filter(Boolean).slice(0, 2).map((x) => x[0].toUpperCase()).join("") || "?");
  const person = (r) => (r.name && r.email && r.name !== r.email ? `${r.name} <${r.email}>` : r.name || r.email || "");

  // ---------------------------------------------------------------- registry (id -> attachment, for the loaded portfolios)
  A.registry = {};
  A.rebuild = function () {
    const reg = {}, st = ATS.store;
    Object.keys(st.portfolios || {}).forEach((pf) => ["kpi", "poap", "raid"].forEach((k) => { const d = st.portfolios[pf][k]; (d && d.attachments || []).forEach((it) => { it.portfolio = pf; it.source = k; reg[it.id] = it; }); }));
    A.registry = reg;
  };
  A.get = (id) => A.registry[id];
  A.forPortfolio = function (pf) { return Object.values(A.registry).filter((x) => x.portfolio === (pf || ATS.store.current)); };

  // ---------------------------------------------------------------- chips
  A.chip = function (ids, label) {
    const list = (ids || []).map(A.get).filter(Boolean);
    if (!list.length) return "";
    const glyphs = list.slice(0, 3).map((it) => (kindOf(it) === "image" && it.size < 120000 ? `<img class="att-th" alt="" src="${dataUrl(it)}">` : `<i class="att-gl">${iconFor(it)}</i>`)).join("");
    const tip = list.map((x) => x.name).join(", ");
    return `<button type="button" class="att-chip" data-att="${list.map((x) => x.id).join(",")}" data-label="${esc(label || "")}" title="${esc(tip)}">${ICON.clip}<span class="att-glyphs">${glyphs}</span><b>${list.length}</b></button>`;
  };
  // vertical list for detail drawers: one button per attachment, all opening the same gallery at that file
  A.list = function (ids, label) {
    const all = (ids || []).filter((id) => A.get(id));
    return '<div class="att-listing">' + all.map((id, i) => { const it = A.get(id), k = kindOf(it); return `<button type="button" class="att-row" data-att="${all.join(",")}" data-start="${i}" data-label="${esc(label || "")}">${k === "image" && it.size < 120000 ? `<img class="att-th" alt="" src="${dataUrl(it)}">` : `<i class="att-gl">${iconFor(it)}</i>`}<span><b>${esc(k === "msg" && it.email && it.email.subject ? it.email.subject : it.name)}</b><small>${esc(k === "msg" && it.email ? [it.email.from && (it.email.from.name || it.email.from.email), when(it.email.date)].filter(Boolean).join(" · ") : KIND[k] + " · " + bytes(it.size))}</small></span></button>`; }).join("") + "</div>";
  };
  A.chipFor = (rec, label) => (rec && rec.att ? A.chip(rec.att, label) : "");
  document.addEventListener("click", (e) => {
    const b = e.target.closest && e.target.closest("[data-att]");
    if (!b) return;
    e.preventDefault(); e.stopPropagation();
    A.open((b.getAttribute("data-att") || "").split(",").filter(Boolean), { label: b.getAttribute("data-label") || "", start: +(b.getAttribute("data-start") || 0) });
  }, true);

  // ---------------------------------------------------------------- download helpers
  A.download = function (name, bytesOrB64, mime) {
    const data = typeof bytesOrB64 === "string" ? A.fromB64(bytesOrB64) : bytesOrB64;
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([data], { type: mime || "application/octet-stream" }));
    a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  };
  const safeName = (n) => String(n || "attachment").replace(/[\\/:*?"<>|]+/g, "_");

  // ---------------------------------------------------------------- viewer
  let V = null; // state: { stack:[{items, i, label, title}], zoom:{...}, mode }
  const el = () => document.getElementById("att-viewer");

  function ensure() {
    if (el()) return el();
    const d = document.createElement("div");
    d.id = "att-viewer"; d.className = "att-viewer"; d.setAttribute("role", "dialog"); d.setAttribute("aria-modal", "true"); d.setAttribute("aria-label", "Attachment viewer");
    d.innerHTML = '<div class="att-back" data-close></div><div class="att-panel"><aside class="att-rail"></aside><section class="att-main"></section></div>';
    document.body.appendChild(d);
    d.addEventListener("click", (e) => {
      if (e.target.closest("[data-close]")) return close();
      const it = e.target.closest("[data-pick]"); if (it) return pick(+it.dataset.pick);
      const nest = e.target.closest("[data-nest]"); if (nest) return openNested(+nest.dataset.nest);
      const dl = e.target.closest("[data-dl]"); if (dl) return doDownload(dl.dataset.dl, dl.dataset.i == null ? null : +dl.dataset.i);
      const tab = e.target.closest("[data-tab]"); if (tab) { V.tab = tab.dataset.tab; return render(); }
      const z = e.target.closest("[data-zoom]"); if (z) return zoom(z.dataset.zoom);
      if (e.target.closest("[data-grid]")) { V.grid = !V.grid; return render(); }
      if (e.target.closest("[data-back]")) { V.stack.pop(); return render(); }
      if (e.target.closest("[data-scope]")) { return scope(e.target.closest("[data-scope]").dataset.scope); }
    });
    d.addEventListener("wheel", (e) => { const s = e.target.closest(".att-stage-img"); if (s) { e.preventDefault(); zoom(e.deltaY < 0 ? "in" : "out", e); } }, { passive: false });
    let drag = null;
    d.addEventListener("pointerdown", (e) => { const s = e.target.closest(".att-stage-img"); if (!s) return; drag = { x: e.clientX - V.z.x, y: e.clientY - V.z.y }; s.setPointerCapture(e.pointerId); s.classList.add("is-drag"); });
    d.addEventListener("pointermove", (e) => { if (!drag) return; V.z.x = e.clientX - drag.x; V.z.y = e.clientY - drag.y; applyZoom(); });
    d.addEventListener("pointerup", () => { drag = null; const s = d.querySelector(".att-stage-img"); if (s) s.classList.remove("is-drag"); });
    d.addEventListener("dblclick", (e) => { if (e.target.closest(".att-stage-img")) zoom(V.z.s > 1.05 ? "fit" : "100"); });
    document.addEventListener("keydown", (e) => {
      if (!V || !el() || !el().classList.contains("open")) return;
      if (e.key === "Escape") { e.stopPropagation(); V.stack.length > 1 ? (V.stack.pop(), render()) : close(); }
      else if (e.key === "ArrowRight" || e.key === "ArrowDown") { e.preventDefault(); step(1); }
      else if (e.key === "ArrowLeft" || e.key === "ArrowUp") { e.preventDefault(); step(-1); }
      else if (e.key === "+" || e.key === "=") zoom("in");
      else if (e.key === "-") zoom("out");
      else if (e.key === "0") zoom("fit");
      else if (e.key === "g" || e.key === "G") { V.grid = !V.grid; render(); }
    }, true);
    return d;
  }

  A.open = function (ids, opts) {
    opts = opts || {};
    const items = (ids || []).map(A.get).filter(Boolean);
    if (!items.length) { ATS.toast && ATS.toast("Attachment not found — reload the data", "err"); return; }
    ensure();
    V = { stack: [{ items, i: Math.min(Math.max(0, opts.start || 0), items.length - 1), label: opts.label || items[0].owner || "", title: opts.label || "Evidence" }], tab: "message", grid: false, z: { s: 1, x: 0, y: 0 }, scopeAll: false, root: items };
    el().classList.add("open"); document.body.classList.add("att-open");
    render();
    const btn = el().querySelector(".att-x"); btn && btn.focus();
  };
  function close() { if (!el()) return; el().classList.remove("open"); document.body.classList.remove("att-open"); const m = el().querySelector(".att-main"); if (m) m.innerHTML = ""; V = null; }
  A.close = close;
  const cur = () => V.stack[V.stack.length - 1];
  const item = () => cur().items[cur().i];
  function pick(i) { cur().i = i; V.tab = "message"; V.grid = false; V.z = { s: 1, x: 0, y: 0 }; render(); }
  function step(d) { const c = cur(), n = c.items.length; if (n < 2) return; c.i = (c.i + d + n) % n; V.tab = "message"; V.z = { s: 1, x: 0, y: 0 }; V.grid = false; render(); }
  function scope(which) {
    const c = V.stack[0];
    if (which === "all") { const all = A.forPortfolio(item().portfolio || ATS.store.current).sort((a, b) => (a.sheet + (1e6 + a.row)).localeCompare(b.sheet + (1e6 + b.row))); c.items = all; c.i = Math.max(0, all.findIndex((x) => x.id === V.root[0].id)); V.scopeAll = true; c.label = "All evidence — " + (ATS.store.current || ""); }
    else { c.items = V.root; c.i = 0; V.scopeAll = false; c.label = V.root[0].owner || ""; }
    V.stack.length = 1; V.grid = false; render();
  }
  function openNested(ai) {
    const parent = item(), em = parent.email; if (!em) return;
    const a = em.attachments[ai]; if (!a) return;
    const sn = A.sniff(a.b64 ? A.fromB64(a.b64) : []), nested = Object.assign({ id: parent.id + "/" + ai, owner: parent.name, nested: true }, a);
    nested.kind = a.email ? "msg" : /^image\//.test(a.mime) || sn.kind === "image" ? "image" : "file";
    if (nested.kind === "image" && !/^image\//.test(nested.mime)) nested.mime = sn.mime;
    if (!nested.email && a.b64 && (sn.kind === "cfb" || /\.msg$/i.test(a.name)) && A.isMsg(A.fromB64(a.b64))) { try { nested.email = A.parseMsg(A.fromB64(a.b64)); nested.kind = "msg"; } catch (e) { /* downloadable file */ } }
    if (nested.kind === "file" && !a.b64) { ATS.toast && ATS.toast("This part of the email has no stored content to open", "err"); return; }
    V.stack.push({ items: [nested], i: 0, label: parent.name, title: parent.name }); V.tab = "message"; V.z = { s: 1, x: 0, y: 0 }; render();
  }
  function doDownload(what, ai) {
    const it = item();
    if (what === "orig") A.download(safeName(it.name), it.b64, it.mime);
    else if (what === "eml") A.download(safeName((it.email.subject || "email") + ".eml"), new TextEncoder().encode(A.toEml(it.email)), "message/rfc822");
    else if (what === "att") { const a = it.email.attachments[ai]; if (a && a.b64) A.download(safeName(a.name), a.b64, a.mime); }
    else if (what === "all") { cur().items.forEach((x, k) => setTimeout(() => A.download(safeName(x.name), x.b64, x.mime), k * 250)); }
  }

  // ---- image zoom / pan
  function applyZoom() { const im = el().querySelector(".att-stage-img img"); if (im) im.style.transform = `translate(${V.z.x}px, ${V.z.y}px) scale(${V.z.s}) rotate(${V.z.r || 0}deg)`; const l = el().querySelector(".att-zoomlbl"); if (l) l.textContent = Math.round(V.z.s * 100) + "%"; }
  function zoom(how, ev) {
    const z = V.z;
    if (how === "fit") V.z = { s: 1, x: 0, y: 0, r: z.r || 0 };
    else if (how === "100") { const im = el().querySelector(".att-stage-img img"); const f = im && im.naturalWidth ? im.naturalWidth / im.clientWidth : 1; V.z = { s: Math.max(1, f), x: 0, y: 0, r: z.r || 0 }; }
    else if (how === "rot") z.r = ((z.r || 0) + 90) % 360;
    else { const f = how === "in" ? 1.25 : 0.8, ns = Math.min(12, Math.max(0.2, z.s * f)); z.x *= ns / z.s; z.y *= ns / z.s; z.s = ns; }
    applyZoom();
  }

  // ---------------------------------------------------------------- rendering
  function railHtml() {
    const c = cur(), top = V.stack.length === 1;
    const rows = c.items.map((it, i) => {
      const k = kindOf(it), th = k === "image" && it.size < 400000 ? `<img src="${dataUrl(it)}" alt="">` : `<span class="att-ico att-ico-${k}">${iconFor(it)}</span>`;
      const sub = k === "msg" && it.email ? esc([it.email.from && (it.email.from.name || it.email.from.email), it.email.date ? new Date(it.email.date).toLocaleDateString("en-GB") : ""].filter(Boolean).join(" · ")) : esc(KIND[k] + " · " + bytes(it.size));
      return `<button type="button" class="att-item${i === c.i ? " is-on" : ""}" data-pick="${i}"><span class="att-thumb">${th}</span><span class="att-meta"><b>${esc(k === "msg" && it.email && it.email.subject ? it.email.subject : it.name)}</b><small>${sub}</small></span></button>`;
    }).join("");
    const back = V.stack.length > 1 ? `<button type="button" class="att-backbtn" data-back>← Back to ${esc(V.stack[V.stack.length - 2].items[V.stack[V.stack.length - 2].i].name)}</button>` : "";
    const scopeBtns = top ? `<div class="att-scope"><button type="button" data-scope="one" class="${V.scopeAll ? "" : "is-on"}">This item</button><button type="button" data-scope="all" class="${V.scopeAll ? "is-on" : ""}">All evidence</button></div>` : "";
    return `<div class="att-rail-h"><div class="att-eyebrow">EVIDENCE</div><div class="att-title" title="${esc(c.label)}">${esc(c.label || "Attachments")}</div><div class="att-count">${c.items.length} file${c.items.length === 1 ? "" : "s"}</div>${back}${scopeBtns}</div><div class="att-list">${rows}</div><div class="att-rail-f"><button type="button" class="att-linkbtn" data-grid>${V.grid ? "Hide" : "Show"} gallery (G)</button>${c.items.length > 1 ? '<button type="button" class="att-linkbtn" data-dl="all">Download all</button>' : ""}</div>`;
  }
  function topbar(it, extra) {
    const k = kindOf(it);
    return `<header class="att-top"><div class="att-top-l"><span class="att-kind att-kind-${k}">${iconFor(it)}${KIND[k]}</span><b class="att-name" title="${esc(it.name)}">${esc(k === "msg" && it.email && it.email.subject ? it.email.subject : it.name)}</b><small>${it.size ? bytes(it.size) : ""}${it.owner && !it.nested ? " · on " + esc(it.owner) : ""}</small></div><div class="att-top-r">${extra || ""}${it.b64 ? `<button type="button" class="att-btn" data-dl="orig" title="Download the original file">${ICON.dl} Download</button>` : ""}<button type="button" class="att-x" data-close aria-label="Close (Esc)">✕</button></div></header>`;
  }
  function gridHtml() {
    return `<div class="att-grid">${cur().items.map((it, i) => { const k = kindOf(it); return `<button type="button" class="att-card" data-pick="${i}"><span class="att-card-th">${k === "image" ? `<img src="${dataUrl(it)}" alt="">` : k === "msg" ? `<span class="att-env">${ICON.mail}<small>${esc((it.email && it.email.subject) || "")}</small></span>` : `<span class="att-ico att-ico-file">${ICON.file}</span>`}</span><b>${esc(it.name)}</b><small>${esc(it.owner || KIND[k])}</small></button>`; }).join("")}</div>`;
  }
  function imageHtml(it) {
    if (it.previewOnly) return `<div class="att-empty"><div class="att-ico att-ico-file">${ICON.file}</div><p>This is a Windows metafile preview that browsers cannot display. Use <b>Download</b> to open the original.</p></div>`;
    return `<div class="att-toolbar"><button type="button" data-zoom="out" title="Zoom out (−)">−</button><span class="att-zoomlbl">100%</span><button type="button" data-zoom="in" title="Zoom in (+)">+</button><button type="button" data-zoom="fit" title="Fit (0)">Fit</button><button type="button" data-zoom="100" title="Actual size">1:1</button><button type="button" data-zoom="rot" title="Rotate">⟳</button><span class="att-hint">scroll to zoom · drag to pan · double-click toggles</span></div><div class="att-stage-img"><img alt="${esc(it.name)}" src="${dataUrl(it)}" draggable="false"></div>`;
  }
  function emailHtml(it) {
    const m = it.email, from = m.from || {}, hu = hue(from.email || from.name || m.subject || "x");
    const chips = (list, lab) => (list.length ? `<div class="att-line"><span>${lab}</span>${list.map((r) => `<em title="${esc(r.email)}">${esc(r.name || r.email)}</em>`).join("")}</div>` : "");
    const atts = (m.attachments || []).filter((a) => !(a.cid && /^image\//.test(a.mime) && m.bodyHtml && m.bodyHtml.includes("data:" + A.safeMime(a.mime)))).map((a, i) => {
      const real = m.attachments.indexOf(a), isImg = /^image\//.test(a.mime);
      return `<span class="att-file">${a.b64 && isImg ? `<img src="data:${A.safeMime(a.mime)};base64,${a.b64}" alt="">` : `<i>${a.embeddedMsg ? ICON.mail : ICON.file}</i>`}<button type="button" data-nest="${real}" title="Open">${esc(a.name)}</button><small>${a.size ? bytes(a.size) : ""}</small>${a.b64 ? `<button type="button" class="att-mini" data-dl="att" data-i="${real}" title="Download">${ICON.dl}</button>` : ""}</span>`;
    }).join("");
    const tab = V.tab || "message";
    // the message HTML is untrusted: no scripts (sandbox without allow-scripts), no network (CSP), links open in a new tab
    const head = `<meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; font-src data:"><base target="_blank"><style>body{font:14px/1.5 -apple-system,Segoe UI,Arial,sans-serif;color:#1d2433;margin:16px;word-wrap:break-word}img{max-width:100%;height:auto}blockquote{border-left:3px solid #dfe3ea;margin-left:0;padding-left:12px;color:#475467}table{max-width:100%}</style>`;
    const html = m.bodyHtml ? "<!doctype html>" + head + m.bodyHtml : "";
    const body = tab === "message"
      ? (html ? `<iframe class="att-mail-frame" sandbox="allow-popups allow-popups-to-escape-sandbox" referrerpolicy="no-referrer" srcdoc="${esc(html)}" title="Email body"></iframe><div class="att-note">Scripts and remote images are blocked for safety.</div>` : `<pre class="att-mail-text">${esc(m.bodyText || "(this email has no body)")}</pre>`)
      : tab === "text" ? `<pre class="att-mail-text">${esc(m.bodyText || "(no plain-text version stored — see the Message tab)")}</pre>`
      : `<pre class="att-mail-text">${esc(m.headers || "(no transport headers stored)")}</pre>`;
    return `<div class="att-mail"><div class="att-mail-head"><span class="att-avatar" style="background:hsl(${hu} 55% 42%)">${esc(initials(from.name || from.email))}</span><div class="att-mail-hd"><h3>${esc(m.subject || "(no subject)")}</h3><div class="att-line"><span>From</span><em title="${esc(from.email)}">${esc(person(from) || "unknown sender")}</em><small>${esc(when(m.date))}</small></div>${chips(m.to, "To")}${chips(m.cc, "Cc")}</div><div class="att-mail-act"><button type="button" class="att-btn" data-dl="eml" title="Open in any mail app (Outlook, Apple Mail, Thunderbird)">Save as .eml</button></div></div>${atts ? `<div class="att-files">${atts}</div>` : ""}<div class="att-tabs"><button type="button" data-tab="message" class="${tab === "message" ? "is-on" : ""}">Message</button><button type="button" data-tab="text" class="${tab === "text" ? "is-on" : ""}">Plain text</button><button type="button" data-tab="headers" class="${tab === "headers" ? "is-on" : ""}">Details</button></div><div class="att-mail-body">${body}</div></div>`;
  }
  function fileHtml(it) {
    if (it.mime === "application/pdf" && it.b64) return `<object class="att-pdf" type="application/pdf" data="${dataUrl(it)}"><div class="att-empty"><div class="att-ico att-ico-file">${ICON.file}</div><p>Your browser cannot show this PDF inline. Use <b>Download</b>.</p></div></object>`;
    return `<div class="att-empty"><div class="att-ico att-ico-file">${ICON.file}</div><h3>${esc(it.name)}</h3><p>${esc(it.mime || "file")} · ${bytes(it.size)}</p><p class="att-dim">A preview is not available for this file type. Download it to open it with your own app.</p><button type="button" class="att-btn att-btn-big" data-dl="orig">${ICON.dl} Download</button></div>`;
  }
  function render() {
    const root = el(); if (!root || !V) return;
    const it = item(); if (!it) return close();
    root.querySelector(".att-rail").innerHTML = railHtml();
    const k = kindOf(it);
    const body = V.grid ? gridHtml() : k === "image" ? imageHtml(it) : k === "msg" && it.email ? emailHtml(it) : fileHtml(it);
    root.querySelector(".att-main").innerHTML = topbar(it) + `<div class="att-body att-body-${V.grid ? "grid" : k}">${body}</div>` + (cur().items.length > 1 && !V.grid ? '<button type="button" class="att-nav att-prev" aria-label="Previous">‹</button><button type="button" class="att-nav att-next" aria-label="Next">›</button>' : "");
    const p = root.querySelector(".att-prev"), n = root.querySelector(".att-next");
    if (p) { p.onclick = () => step(-1); n.onclick = () => step(1); }
    if (k === "image" && !V.grid) applyZoom();
  }

  // problems to show on the Data Quality page
  A.quality = function () {
    const w = [], st = ATS.store;
    ["kpi", "poap", "raid"].forEach((k) => { const d = st[k]; if (!d) return; (d.attachmentWarnings || []).forEach((m) => w.push({ sheet: "Attachments", msg: m })); (d.attachments || []).filter((x) => !x.linked).forEach((x) => w.push({ sheet: x.sheet, msg: `"${x.name}" (row ${x.row}) is not on a data row, so it appears only in the Evidence library. Move it onto the row of the note, defect, risk or plan item it belongs to.` })); });
    return w;
  };

  // ---------------------------------------------------------------- Evidence library page (Appendix of Monthly Council / Weekly Report)
  ATS.pages = ATS.pages || {};
  ATS.pages.evidence = function (env) {
    const U = ATS.ui, items = A.forPortfolio(ATS.store.current);
    const prefix = env.prefix, state = (ATS.pages._evState = ATS.pages._evState || { q: "", kind: "all" });
    const foot = `Source: attachments inside the Excel workbooks · ${esc(env.data.config.programName)}`;
    if (!items.length) {
      return { html: U.slide({ id: prefix + "slide-evidence", eyebrow: "Appendix", title: "Evidence library", sub: "Pictures and Outlook emails attached in the workbooks", single: true, foot,
        left: `<div class="sx-callout info"><b>No attachments found in this portfolio's workbooks.</b><br>In Excel, put a picture on a note's row with <i>Insert → Pictures</i>, or an email with <i>Insert → Object → Create from file</i> (tick <i>Display as icon</i> and pick the .msg), or drag an email from Outlook onto the row. Save, then click <b>Refresh data</b>. Each attachment shows next to its note, defect, risk or plan item and is collected here.</div>` }), draw() {} };
    }
    const counts = { all: items.length, image: items.filter((x) => kindOf(x) === "image").length, msg: items.filter((x) => kindOf(x) === "msg").length, file: items.filter((x) => kindOf(x) === "file").length };
    const rows = () => {
      const q = state.q.toLowerCase(), list = items.filter((x) => (state.kind === "all" || kindOf(x) === state.kind) && (!q || (x.name + " " + (x.owner || "") + " " + x.sheet + " " + (x.email ? x.email.subject + " " + (x.email.from && x.email.from.name) + " " + x.email.bodyText : "")).toLowerCase().includes(q)));
      const groups = {}; list.forEach((x) => (groups[x.sheet] = groups[x.sheet] || []).push(x));
      if (!list.length) return '<div class="sx-empty-note">Nothing matches.</div>';
      return Object.keys(groups).map((sh) => `<div class="att-lib-g"><div class="sx-ctitle">${esc(sh.replace(/_/g, " "))} <span class="sx-csub">· ${groups[sh].length}</span></div><div class="att-lib-grid">${groups[sh].map((x) => { const k = kindOf(x); return `<button type="button" class="att-card" data-att="${x.id}" data-label="${esc(x.owner || "")}"><span class="att-card-th">${k === "image" ? `<img src="${dataUrl(x)}" alt="">` : k === "msg" ? `<span class="att-env">${ICON.mail}<small>${esc((x.email && x.email.subject) || "")}</small></span>` : `<span class="att-ico att-ico-file">${ICON.file}</span>`}</span><b>${esc(x.name)}</b><small>${esc(x.owner || "")}</small></button>`; }).join("")}</div></div>`).join("");
    };
    const left = `<div class="sx-stats">${U.stat(counts.all, "Attachments")}${U.stat(counts.image, "Pictures")}${U.stat(counts.msg, "Outlook emails")}${U.stat(counts.file, "Other files")}</div>
      <div class="att-lib-bar"><div class="att-seg" id="${prefix}ev-kind">${[["all", "All"], ["image", "Pictures"], ["msg", "Emails"], ["file", "Files"]].map(([k, l]) => `<button type="button" data-k="${k}" class="${state.kind === k ? "is-on" : ""}">${l} <small>${counts[k]}</small></button>`).join("")}</div><input type="search" id="${prefix}ev-q" placeholder="Search names, subjects, senders, notes…" value="${esc(state.q)}"><button type="button" class="btn btn-outline btn-sm" data-att="${items.map((x) => x.id).join(",")}" data-label="All evidence">Open as gallery</button></div>
      <div id="${prefix}ev-list">${rows()}</div>`;
    return { html: U.slide({ id: prefix + "slide-evidence", eyebrow: "Appendix", title: "Evidence library", sub: "Everything attached in the workbooks, collected in one place — click to open", left, single: true, foot }),
      draw(root) {
        const list = root.querySelector("#" + prefix + "ev-list"); if (!list) return;
        root.querySelectorAll("#" + prefix + "ev-kind button").forEach((b) => (b.onclick = () => { state.kind = b.dataset.k; root.querySelectorAll("#" + prefix + "ev-kind button").forEach((x) => x.classList.toggle("is-on", x === b)); list.innerHTML = rows(); }));
        const q = root.querySelector("#" + prefix + "ev-q"); q.oninput = () => { state.q = q.value; list.innerHTML = rows(); };
      } };
  };
})();
