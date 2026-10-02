/* ============================================================
   PDF export — renders the dashboard pages to a PDF (jsPDF + html2canvas, all bundled, works offline)
   Monthly Council / Weekly Report : one PDF page per report page (cover + every page of the pack)
   Programme Overview / POAP       : the current view as a single page
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const D = ATS.date;
  const NAVY = [15, 31, 61], GREY = [107, 114, 128];
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const frame = () => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  const slug = (s) => String(s || "").replace(/[^A-Za-z0-9]+/g, "_").replace(/^_|_$/g, "");

  // capture one element as a JPEG data-url + pixel size (animations/transitions frozen so nothing is half-faded)
  async function snap(el) {
    const st = document.createElement("style");
    st.textContent = "*{animation:none!important;transition:none!important}.sx-foot-btns,.poap-noprint{display:none!important}";
    document.head.appendChild(st);
    try {
      const c = await window.html2canvas(el, { backgroundColor: "#ffffff", scale: 1.8, useCORS: true, logging: false, windowWidth: Math.max(1600, document.documentElement.clientWidth) });
      return { img: c.toDataURL("image/jpeg", 0.92), w: c.width, h: c.height };
    } finally { st.remove(); }
  }

  function newDoc() {
    const J = window.jspdf && window.jspdf.jsPDF;
    if (!J) throw new Error("PDF library missing");
    return new J({ orientation: "l", unit: "mm", format: "a4", compress: true });
  }

  const PW = 297, M = 8, HEAD = 13, FOOT = 8;
  function addImagePage(doc, first, shot, head) {
    const iw = PW - 2 * M, ih = (shot.h / shot.w) * iw;
    const ph = Math.max(210, HEAD + ih + FOOT + 4);
    if (first) { doc.deletePage(1); }
    doc.addPage([Math.min(PW, ph), Math.max(PW, ph)], ph > PW ? "p" : "l");
    doc.setFillColor(...NAVY); doc.rect(0, 0, PW, HEAD - 3, "F");
    doc.setTextColor(255, 255, 255); doc.setFont("helvetica", "bold"); doc.setFontSize(10);
    doc.text(head.left, M, 6.2);
    doc.setFont("helvetica", "normal"); doc.setFontSize(9);
    doc.text(head.right || "", PW - M, 6.2, { align: "right" });
    doc.addImage(shot.img, "JPEG", M, HEAD, iw, ih, undefined, "FAST");
    return ph;
  }
  function footers(doc, note) {
    const n = doc.getNumberOfPages();
    for (let i = 1; i <= n; i++) {
      doc.setPage(i);
      const h = doc.internal.pageSize.getHeight();
      doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(...GREY);
      doc.text(note, M, h - 4);
      doc.text(`Page ${i} of ${n}`, PW - M, h - 4, { align: "right" });
    }
  }
  function cover(doc, lines) {
    doc.setFillColor(...NAVY); doc.rect(0, 0, 297, 210, "F");
    doc.setFillColor(0, 164, 153); doc.rect(0, 0, 8, 210, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold"); doc.setFontSize(34); doc.text(lines.title, 24, 84);
    doc.setFont("helvetica", "normal"); doc.setFontSize(20); doc.text(lines.sub, 24, 100);
    doc.setFontSize(14); doc.setTextColor(202, 220, 252); doc.text(lines.period, 24, 116);
    doc.setFontSize(10); doc.text(lines.foot, 24, 192);
  }
  const stamp = () => "Generated " + new Date().toLocaleDateString("en-GB", { day: "numeric", month: "long", year: "numeric" }) + (ATS.isSample() ? " · SAMPLE DATA" : "");

  async function sectionPdf(sec, kind, view) {
    const doc = newDoc();
    const env0 = sec._env;
    if (!env0 || env0.data !== ATS.store.kpi) throw new Error("this portfolio has no weekly data yet, so there is nothing to export");
    const pages = env0.pages, keep = env0.state.page;
    const who = ATS.store.current || (env0.data.config && env0.data.config.programName) || "";
    const title = kind === "month" ? "Monthly Council" : "Weekly Report";
    const periodLabel = kind === "month" ? D.fmtMonth(env0.state.key) : "Week ending " + D.fmt(env0.state.key);
    cover(doc, { title: who || "ATS Programme", sub: title, period: periodLabel, foot: stamp() });
    let first = false;
    Chart.defaults.animation = false;
    try {
      for (let i = 0; i < pages.length; i++) {
        ATS.toast(`Rendering page ${i + 1} of ${pages.length}…`);
        sec.go(pages[i].id);
        await frame(); await sleep(150);
        const stage = document.querySelector("#" + sec._env.prefix + "stage");
        const shot = await snap(stage);
        addImagePage(doc, first, shot, { left: `${who}  ·  ${title}  ·  ${periodLabel}`, right: `${pages[i].group} — ${pages[i].label}` });
      }
    } finally {
      Chart.defaults.animation = { duration: 350 };
      sec.go(keep);
    }
    footers(doc, `${who} · ${title} · ${periodLabel}  ·  ${stamp()}`);
    const name = (kind === "month" ? "Monthly_Council_" : "Weekly_Report_") + (slug(who) ? slug(who) + "_" : "") + env0.state.key;
    doc.save(name + ".pdf");
    return name + ".pdf";
  }

  async function singleViewPdf(el, label, fileBase) {
    const doc = newDoc();
    Chart.defaults.animation = false;
    let shot;
    try { await frame(); shot = await snap(el); } finally { Chart.defaults.animation = { duration: 350 }; }
    addImagePage(doc, true, shot, { left: `${label}${ATS.store.current && label !== "Programme Overview" ? "  ·  " + ATS.store.current : ""}`, right: new Date().toLocaleDateString("en-GB") });
    footers(doc, stamp());
    doc.save(fileBase + ".pdf");
    return fileBase + ".pdf";
  }

  ATS.exportPdf = async function () {
    if (!window.html2canvas) { ATS.toast("Image library missing — cannot build PDF", "err"); return; }
    const view = ATS.shell.currentView();
    try {
      ATS.toast("Building PDF…");
      let saved;
      if (view === "monthly" || view === "weekly") {
        const sec = view === "monthly" ? ATS.monthly : ATS.weekly;
        saved = await sectionPdf(sec, view === "monthly" ? "month" : "week", view);
      } else if (view === "poap") {
        const root = document.querySelector("#poap-root");
        if (!root) { ATS.toast("Nothing to export on this page", "err"); return; }
        const done = ATS.poap && ATS.poap.prepareExport ? ATS.poap.prepareExport() : null;
        let target = (done && done.el) || root;
        try { saved = await singleViewPdf(target, "POAP — Plan on a Page", "POAP_" + (slug(ATS.store.current) || "Portfolio") + "_" + new Date().toISOString().slice(0, 10)); }
        finally { if (done && done.restore) done.restore(); }
      } else {
        const el = document.querySelector("#view-" + view);
        const last = ATS.programme && ATS.programme.last && ATS.programme.last();
        saved = await singleViewPdf(el, "Programme Overview", "Programme_Overview_" + ((last && last.period && last.period.key) || new Date().toISOString().slice(0, 10)));
      }
      ATS.toast("Saved " + saved, "ok");
    } catch (e) { console.error(e); ATS.toast("PDF export failed: " + ((e && e.message) || "unexpected error — see the browser console"), "err"); }
  };
})();
