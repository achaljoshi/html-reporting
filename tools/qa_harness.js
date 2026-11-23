/* QA harness — paste/inject into the running dashboard to sweep every page for obvious defects.
   usage (browser console or automation):  fetch("tools/qa_harness.js").then(r => r.text()).then(eval)
   then:  await QA.sweep("Self Assessment")      -> list of problems
          await QA.sweepAll()                    -> every portfolio
   Checks: JS errors, "failed to render" callouts, undefined/NaN/Infinity/[object] text, horizontal overflow,
           zero-size charts, empty stages, missing page titles. */
(function () {
  const QA = (window.QA = window.QA || {});
  QA.errors = QA.errors || [];
  if (!QA._hooked) {
    QA._hooked = true;
    window.addEventListener("error", (e) => QA.errors.push("error: " + e.message));
    window.addEventListener("unhandledrejection", (e) => QA.errors.push("rejection: " + (e.reason && e.reason.message || e.reason)));
    const ce = console.error; console.error = function () { QA.errors.push("console.error: " + Array.from(arguments).map(String).join(" ").slice(0, 200)); ce.apply(console, arguments); };
    window.requestAnimationFrame = (f) => setTimeout(() => f(performance.now()), 16); // hidden panes do not fire rAF
  }
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const BAD = /\b(undefined|NaN|Infinity)\b|\[object|\bnull\b(?!\s*[,)])/;

  QA.checked = 0;
  function inspect(where, root, out) {
    QA.checked++;
    if (!root) { out.push(where + ": no root element"); return; }
    const txt = root.innerText || "";
    const m = BAD.exec(txt);
    if (m) out.push(where + ': suspicious text "' + txt.slice(Math.max(0, m.index - 30), m.index + 30).replace(/\s+/g, " ") + '"');
    if (/failed to render|could not be built|could not be drawn/i.test(txt)) out.push(where + ": page failed to render");
    if (document.documentElement.scrollWidth > window.innerWidth + 2) out.push(where + ": page scrolls horizontally (" + document.documentElement.scrollWidth + " > " + window.innerWidth + ")");
    if (txt.trim().length < 20) out.push(where + ": page is empty");
    const rb = root.getBoundingClientRect();
    const wide = Array.from(root.querySelectorAll("canvas,.sx-panel,.sx-stat,img,svg")).filter((e) => { const r = e.getBoundingClientRect(); return r.width > 0 && r.right > rb.right + 3; });
    if (wide.length) out.push(where + ": " + wide.length + " element(s) overflow the page width (" + wide[0].className + ")");
    Array.from(root.querySelectorAll("canvas")).forEach((c) => { if (c.offsetParent && (c.clientWidth < 20 || c.clientHeight < 20)) out.push(where + ": chart canvas has no size"); });
    Array.from(root.querySelectorAll("table")).forEach((t) => { if (t.querySelectorAll("tbody tr").length === 0) out.push(where + ": table without rows"); });
  }

  QA.sweepSection = async function (pf, view, nPeriods) {
    const out = [], sec = view === "monthly" ? ATS.monthly : ATS.weekly;
    ATS.setPortfolio(pf); ATS.shell.go(view); await sleep(250);
    const env0 = sec._env; if (!env0) { out.push(pf + "/" + view + ": section did not render"); return out; }
    const keys = (view === "monthly" ? ATS.kpi.months(env0.data) : ATS.kpi.weeks(env0.data)).slice(-(nPeriods || 2));
    for (const k of keys) {
      sec.state.key = k; sec.render(); await sleep(150);
      for (const p of sec._env.pages) {
        sec.go(p.id); await sleep(110);
        const where = pf + "/" + view + "/" + k + "/" + p.id;
        const stage = document.querySelector("#" + sec._env.prefix + "stage");
        inspect(where, stage, out);
      }
    }
    return out;
  };
  QA.sweep = async function (pf, nPeriods) {
    const before = QA.errors.length;
    const out = [].concat(await QA.sweepSection(pf, "monthly", nPeriods), await QA.sweepSection(pf, "weekly", nPeriods));
    QA.errors.slice(before).forEach((e) => out.push(pf + ": " + e));
    return out;
  };
  QA.sweepAll = async function () { let out = []; for (const pf of ATS.portfolioNames()) out = out.concat(await QA.sweep(pf, 2)); return out; };

  // ---- workbook helpers for crafting edge-case data in the browser
  QA.getWB = async (url) => XLSX.read(new Uint8Array(await fetch(url).then((r) => r.arrayBuffer())), { type: "array", cellDates: true });
  QA.toFile = (wb, name, dir) => { const f = new File([XLSX.write(wb, { type: "array", bookType: "xlsx" })], name, { lastModified: Date.now() }); f._path = dir ? dir + "/" + name : name; return f; };
  QA.mutate = (wb, name, fn) => { const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: "" }); fn(aoa); wb.Sheets[name] = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true }); };
  QA.variant = async (pf, fn, base) => {
    const wb = await QA.getWB(base || "samples/ATS_Program_SAMPLE/PAYE/ATS_Weekly_Data.xlsx");
    QA.mutate(wb, "Config", (a) => { for (const r of a) if (String(r[0]).trim() === "Portfolio Name") r[1] = pf; });
    await fn(wb); return QA.toFile(wb, "ATS_Weekly_Data - " + pf + ".xlsx", pf);
  };
  QA.clear = (name) => (wb) => QA.mutate(wb, name, (a) => { a.length = 5; });
})();
