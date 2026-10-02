/* ============================================================
   App shell — sidebar, routing between sections, portfolio selector, data loading UI
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const $ = ATS.dom.$, $all = ATS.dom.$all, esc = ATS.dom.esc;

  const SECTIONS = {
    programme: { title: "Programme Overview", crumb: "All portfolios side by side — click any dot, card or bar to drill in", mod: () => ATS.programme, portfolioPicker: false },
    monthly: { title: "Monthly Council", crumb: "Unified quality council pack — arrow keys move between pages", mod: () => ATS.monthly, portfolioPicker: true },
    weekly: { title: "Weekly Report", crumb: "Weekly delivery & quality report — same data, sliced by week", mod: () => ATS.weekly, portfolioPicker: true },
    poap: { title: "POAP — Plan on a Page", crumb: "Interactive programme roadmap for the selected portfolio", portfolioPicker: true },
  };
  let current = "programme";
  let poapMounted = false, poapFor = null;

  const shell = (ATS.shell = {});

  shell.go = function (id) {
    if (!SECTIONS[id]) id = "programme";
    current = id;
    try { sessionStorage.setItem("ats_last_view", id); } catch (e) {}
    $all(".nav-item").forEach((b) => b.classList.toggle("active", b.getAttribute("data-view") === id));
    $all(".view").forEach((v) => v.classList.remove("active"));
    $("#view-" + id).classList.add("active");
    const s = SECTIONS[id];
    $("#view-title").textContent = s.title;
    $("#view-crumb").textContent = s.crumb;
    shell.render();
    window.scrollTo({ top: 0, behavior: "auto" });
  };

  shell.render = function () {
    updateChrome();
    const el = $("#view-" + current);
    if (current === "poap") return renderPoap(el);
    const mod = SECTIONS[current].mod();
    if (!mod) return;
    if (!mod.root()) mod.mount(el);
    mod.render();
  };

  function renderPoap(el) {
    const st = ATS.store;
    if (!ATS.poap || !ATS.poap.render) { el.innerHTML = `<div class="card card-pad sx-load-card"><h2>POAP module not found</h2><p class="text-sub">assets/js/poap.js is missing.</p></div>`; poapMounted = false; return; }
    if (!ATS.hasData()) { el.innerHTML = ATS.emptyCard("POAP — Plan on a Page"); ATS.wireEmptyCard(el); poapMounted = false; return; }
    if (!st.poap) {
      el.innerHTML = `<div class="card card-pad sx-load-card"><div class="sx-eyebrow">POAP — Plan on a Page</div><h2 style="margin:6px 0 8px">No POAP for ${esc(st.current)}</h2>
        <p class="text-sm text-sub" style="line-height:1.6;margin:0">Put this portfolio's <b>ATS_POAP.xlsx</b> (copy of <code>templates/ATS_POAP_Template.xlsx</code>) in the same folder as its weekly workbook, then click <b>Refresh data</b>.</p></div>`;
      poapMounted = false; return;
    }
    if (!poapMounted || poapFor !== st.current || !el.querySelector("#poap-root")) {
      el.innerHTML = '<div id="poap-root"></div>';
      ATS.poap.mount(el.querySelector("#poap-root"));
      poapMounted = true; poapFor = st.current;
    }
    ATS.poap.render(st.poap, { today: new Date(), sampleMode: ATS.isSample(), kpi: st.kpi, raid: st.raid });
  }

  // ------------------------------------------------------------------ chrome: selector + sidebar data panel
  function updateChrome() {
    const st = ATS.store, meta = st.meta || {};
    document.body.classList.toggle("is-sample", !!ATS.isSample());
    const names = ATS.portfolioNames();
    // portfolio picker (only for sections that show ONE portfolio)
    const wrap = $("#portfolio-wrap"), sel = $("#portfolio-select");
    const show = names.length > 0 && SECTIONS[current].portfolioPicker;
    wrap.style.display = show ? "inline-flex" : "none";
    if (show) {
      sel.innerHTML = names.map((n) => `<option value="${esc(n)}" ${n === st.current ? "selected" : ""}>${esc(n)}</option>`).join("");
      sel.disabled = names.length < 2;
    }
    // sidebar
    const k = st.kpi;
    $("#side-program").textContent = (k && k.config.programme) || (names.length > 1 ? "Programme" : (k && k.config.programName) || "Executive Dashboard");
    const files = [["kpi", "Weekly data workbook"], ["poap", "POAP workbook"], ["raid", "RAID log"]];
    const head = names.length ? `<div class="side-file ok" style="color:#fff;font-weight:700;margin-bottom:2px">${names.length} portfolio${names.length === 1 ? "" : "s"} loaded</div>${st.current && current !== "programme" ? `<div class="side-file" style="margin-bottom:4px">Showing: <b style="color:#dbe6ff">${esc(st.current)}</b></div>` : ""}` : "";
    $("#side-files").innerHTML = head + (names.length && current !== "programme" ? files.map(([key, label]) => `<div class="side-file ${st[key] ? "ok" : ""}"><i></i>${esc(label)}</div>`).join("") : "");
    const when = meta.loadedAt ? new Date(meta.loadedAt).toLocaleString() : "";
    $("#program-status").innerHTML = meta.source === "sample" ? "<b>Sample data</b> loaded — load your own folder to replace it." : meta.source === "folder" ? `Folder <b>${esc(meta.label || "")}</b><br>Loaded ${esc(when)}` : "No data loaded yet.";
  }

  // ------------------------------------------------------------------ loading
  function summarise(report) {
    return report.map((r) => `${r.kind === "ignored" ? "–" : r.kind === "error" ? "✗" : "✓"} ${r.dir ? r.dir + "/" : ""}${r.name}${r.note ? " — " + r.note : ""}`).join("\n");
  }
  shell.load = async function (kind) {
    if (kind === "sample") { ATS.useSample(); ATS.toast("Sample data loaded", "ok"); return; }
    ATS.toast("Reading workbooks…");
    const res = kind === "files" ? await ATS.pickFiles() : kind === "refresh" ? await ATS.resync() : await ATS.pickFolder();
    if (res.cancelled) return;
    if (res.denied) { ATS.toast("Folder access was not granted", "err"); return; }
    if (!res.ok) {
      alert("No ATS workbooks were recognised in that selection.\n\nExpected files containing:\n • sheets Weekly_Snapshot + Defect_Log (weekly data workbook — one per portfolio)\n • sheet POAP_Plan (POAP workbook)\n • sheets Risks + Issues (RAID log)\n\n" + summarise(res.report));
      return;
    }
    const n = (res.portfolios || []).length;
    ATS.toast(`Loaded ${n} portfolio${n === 1 ? "" : "s"}: ${res.portfolios.slice(0, 4).join(", ")}${n > 4 ? "…" : ""}`, "ok");
  };

  // ------------------------------------------------------------------ boot
  async function boot() {
    $("#btn-load-program").onclick = () => shell.load("folder");
    $("#btn-refresh-program").onclick = () => shell.load("refresh");
    $("#portfolio-select").onchange = (e) => ATS.setPortfolio(e.target.value);
    $("#btn-present").onclick = () => { document.body.classList.add("present-mode"); $("#btn-present").style.display = "none"; $("#btn-print").style.display = "none"; $("#btn-present-exit").style.display = "inline-flex"; };
    $("#btn-present-exit").onclick = () => { document.body.classList.remove("present-mode"); $("#btn-present").style.display = ""; $("#btn-print").style.display = ""; $("#btn-present-exit").style.display = "none"; };
    $("#btn-print").onclick = () => window.print();
    ATS.on("data", () => { poapMounted = false; shell.render(); });
    ATS.on("portfolio", () => { poapMounted = false; shell.render(); });

    const had = await ATS.loadPersisted();
    if (!had || ATS.store.meta.source === "sample") ATS.useSample();
    let start = "programme";
    try { start = sessionStorage.getItem("ats_last_view") || "programme"; } catch (e) {}
    if (ATS.portfolioNames().length === 1 && start === "programme") start = "monthly";
    $all(".nav-item").forEach((b) => (b.onclick = () => shell.go(b.getAttribute("data-view"))));
    shell.go(start);

    document.addEventListener("keydown", (e) => {
      if (e.target && /^(INPUT|SELECT|TEXTAREA)$/.test(e.target.tagName)) return;
      if (e.key === "Escape") { const b = $("#btn-present-exit"); if (b && b.offsetParent) b.click(); }
      if (current === "monthly" || current === "weekly") {
        const mod = SECTIONS[current].mod();
        if (e.key === "ArrowRight" || e.key === "PageDown") mod.step(1);
        else if (e.key === "ArrowLeft" || e.key === "PageUp") mod.step(-1);
      }
    });
  }
  document.addEventListener("DOMContentLoaded", boot);
})();
