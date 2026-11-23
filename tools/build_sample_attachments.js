#!/usr/bin/env node
/* Adds FICTIONAL demo attachments (pictures + Outlook emails) to the "Self Assessment" sample workbooks, so the Evidence viewer has something to show.
   Pictures are drawn here from SVG; emails are built with tools/make_msg.js. Nothing real is involved.
   usage:  NODE_PATH=$(npm root -g) node tools/build_sample_attachments.js      (then: node tools/build_sample_js.js)
   Safe to re-run: files that already contain attachments are skipped. */
"use strict";
const fs = require("fs"), path = require("path");
const root = path.resolve(__dirname, "..");
global.window = global;
global.XLSX = require(path.join(root, "assets/lib/xlsx.full.min.js"));
["core.js", "kpi-data.js", "raid.js", "attachments.js"].forEach((f) => require(path.join(root, "assets/js", f)));
require(path.join(root, "assets/js/poap.js"));
const sharp = require("sharp");
const { embed } = require("./xlsx_embed.js");
const { makeMsg } = require("./make_msg.js");
const ATS = global.ATS;

const svg = (w, h, body) => `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" font-family="Helvetica,Arial,sans-serif">${body}</svg>`;
const png = (s) => sharp(Buffer.from(s)).png({ compressionLevel: 9, palette: true }).toBuffer();
const IMG = {
  env: svg(720, 380, `<rect width="720" height="380" fill="#0f1c3f"/><text x="28" y="40" fill="#fff" font-size="20" font-weight="700">Integration environment — health (last 7 days)</text><text x="28" y="62" fill="#9fb4e8" font-size="13">Fictional demonstration data</text>
    <g stroke="#27407a"><path d="M40 300H690M40 240H690M40 180H690M40 120H690"/></g><g fill="#9fb4e8" font-size="11"><text x="8" y="304">0%</text><text x="8" y="244">50%</text><text x="4" y="124">100%</text></g>
    <rect x="360" y="100" width="150" height="200" fill="#ef4444" opacity=".25"/><text x="372" y="124" fill="#fecaca" font-size="12" font-weight="700">OUTAGE  09:10 – 16:40</text>
    <polyline fill="none" stroke="#22c55e" stroke-width="3" points="40,110 120,118 200,112 280,120 350,116 360,296 510,296 520,122 600,114 690,112"/>
    <g fill="#cbd5e1" font-size="12"><text x="40" y="330">Mon</text><text x="150" y="330">Tue</text><text x="260" y="330">Wed</text><text x="370" y="330">Thu</text><text x="480" y="330">Fri</text><text x="590" y="330">Sat</text></g>
    <rect x="28" y="344" width="10" height="10" fill="#22c55e"/><text x="44" y="354" fill="#e2e8f0" font-size="12">Availability</text>`),
  err: svg(640, 360, `<rect width="640" height="360" fill="#eef2f7"/><rect x="60" y="40" width="520" height="280" rx="12" fill="#fff" stroke="#cbd5e1"/><rect x="60" y="40" width="520" height="44" rx="12" fill="#1b2a4a"/><text x="84" y="68" fill="#fff" font-size="16" font-weight="700">Self Assessment — Submit return</text>
    <rect x="84" y="108" width="472" height="64" rx="8" fill="#fef3f2" stroke="#f6cdc8"/><text x="104" y="136" fill="#b42318" font-size="15" font-weight="700">We could not submit your return</text><text x="104" y="158" fill="#7a271a" font-size="13">Upstream service timed out after 30 s (HTTP 504). Reference SA-106.</text>
    <rect x="84" y="196" width="220" height="28" rx="6" fill="#e5e7eb"/><rect x="84" y="236" width="300" height="28" rx="6" fill="#e5e7eb"/><rect x="420" y="268" width="136" height="36" rx="8" fill="#0ea5a0"/><text x="446" y="291" fill="#fff" font-size="14" font-weight="700">Try again</text>`),
  burn: svg(720, 380, `<rect width="720" height="380" fill="#fff"/><text x="28" y="38" fill="#101828" font-size="19" font-weight="700">SIT execution burn-up</text><text x="28" y="58" fill="#667085" font-size="12">Fictional demonstration data</text>
    <g stroke="#e4e7ec"><path d="M50 320H690M50 260H690M50 200H690M50 140H690M50 80H690"/></g>
    <polyline fill="none" stroke="#98a2b3" stroke-width="3" stroke-dasharray="8 6" points="50,310 690,90"/>
    <polyline fill="none" stroke="#0ea5a0" stroke-width="4" points="50,312 130,300 210,282 290,262 370,226 450,196 530,176"/>
    <circle cx="530" cy="176" r="6" fill="#0ea5a0"/><text x="540" y="172" fill="#0a7c78" font-size="12" font-weight="700">199 executed</text>
    <rect x="50" y="344" width="10" height="10" fill="#0ea5a0"/><text x="66" y="354" fill="#475467" font-size="12">Executed</text><rect x="150" y="344" width="10" height="3" fill="#98a2b3"/><text x="166" y="354" fill="#475467" font-size="12">Plan</text>`),
};

const P = { alex: { name: "Alex Morgan", email: "alex.morgan@example.com" }, priya: { name: "Priya Nair", email: "priya.nair@example.com" }, sam: { name: "Sam Taylor", email: "sam.taylor@example.com" },
  jo: { name: "Jo Fletcher", email: "jo.fletcher@example.com" }, tm: { name: "Test Manager", email: "test.manager@example.com" }, env: { name: "Environment Team", email: "environments@example.com" } };
const html = (body) => `<html><body style="font-family:Calibri,Arial,sans-serif;font-size:14px">${body}<p style="color:#667085;font-size:12px">This is fictional demonstration content.</p></body></html>`;

(async () => {
  const pic = { env: await png(IMG.env), err: await png(IMG.err), burn: await png(IMG.burn) };
  const MAIL = {
    outage: makeMsg({ subject: "RE: Integration environment outage — SIT blocked", from: P.env, to: [P.tm], cc: [P.alex], date: "2026-09-24T16:55:00Z",
      body: "Hi,\n\nThe integration environment was unavailable from 09:10 to 16:40 today while the gateway pool was rebuilt. SIT execution was blocked for the whole period.\n\nRoot cause: certificate rotation failed on node 2. A permanent fix is scheduled for Tuesday.\n\nRegards,\nEnvironment Team",
      html: html("<p>Hi,</p><p>The integration environment was <b>unavailable from 09:10 to 16:40</b> today while the gateway pool was rebuilt. SIT execution was blocked for the whole period.</p><table border='1' cellpadding='6' style='border-collapse:collapse'><tr><th>Item</th><th>Detail</th></tr><tr><td>Root cause</td><td>Certificate rotation failed on node 2</td></tr><tr><td>Permanent fix</td><td>Tuesday</td></tr></table><p>Regards,<br>Environment Team</p>"),
      attachments: [{ name: "environment-health.png", mime: "image/png", bytes: pic.env }] }),
    vendor: makeMsg({ subject: "Vendor build 3.2 — delivery date moving", from: P.jo, to: [P.tm, P.priya], date: "2026-09-29T08:12:00Z",
      body: "Team,\n\nBuild 3.2 will land on 14 October instead of 9 October because of a late defect fix in the payments interface. We will confirm entry criteria on Monday.\n\nJo",
      html: html("<p>Team,</p><p>Build 3.2 will land on <b>14 October</b> instead of 9 October because of a late defect fix in the payments interface. We will confirm entry criteria on Monday.</p><p>Jo</p>") }),
    decision: makeMsg({ subject: "Approved: move Cycle 2 start to align with the vendor drop", from: P.sam, to: [P.tm], cc: [P.alex, P.priya], date: "2026-09-29T15:40:00Z",
      body: "Steering group confirms the decision to move the Cycle 2 start by one week. Please update the plan and the RAID log.\n\nSam (Chair)",
      html: html("<p>Steering group <b>confirms</b> the decision to move the Cycle 2 start by one week. Please update the plan and the RAID log.</p><p>Sam (Chair)</p>") }),
    defect: makeMsg({ subject: "SA-106 — submission timeout, evidence attached", from: P.priya, to: [P.tm], cc: [P.alex], date: "2026-09-03T10:20:00Z",
      body: "Reproduced three times in SIT. The submit call times out after 30 s. Screenshot attached. Gateway team have the trace id.\n\nPriya",
      html: html("<p>Reproduced three times in SIT. The submit call times out after <b>30 s</b>. Screenshot attached. Gateway team have the trace id.</p><p>Priya</p>"),
      attachments: [{ name: "SA-106-error.png", mime: "image/png", bytes: pic.err }] }),
    tsr: makeMsg({ subject: "TSR response — impact assessment returned", from: P.alex, to: [P.tm], date: "2026-10-01T11:05:00Z",
      body: "Impact assessment for the open TSR was returned in 3 working days, inside the 10-day limit. No change to test scope.\n\nAlex",
      html: html("<p>Impact assessment for the open TSR was returned in <b>3 working days</b>, inside the 10-day limit. No change to test scope.</p><p>Alex</p>") }),
    plan: makeMsg({ subject: "Release 1 SIT start — entry criteria agreed", from: P.alex, to: [P.tm, P.sam], date: "2026-07-21T09:00:00Z",
      body: "Entry criteria for the Release 1 SIT start are agreed: environment smoke test green, masked data loaded, build deployed.\n\nAlex",
      html: html("<p>Entry criteria for the Release 1 SIT start are agreed:</p><ul><li>environment smoke test green</li><li>masked data loaded</li><li>build deployed</li></ul><p>Alex</p>") }),
  };

  const dir = path.join(root, "samples", "ATS_Program_SAMPLE", "Self Assessment");
  const open = (f) => XLSX.read(fs.readFileSync(path.join(dir, f)), { type: "buffer", cellDates: true });
  const lastCol = (wb, sheet) => { const ref = XLSX.utils.decode_range(wb.Sheets[sheet]["!ref"]); return ref.e.c + 2; };   // first free column to the right of the data (1-based)

  async function apply(file, plan) {
    const full = path.join(dir, file), buf = fs.readFileSync(full);
    if (ATS.att.extract(new Uint8Array(buf)).items.length) { console.log("skip (already has attachments):", file); return; }
    const wb = open(file), items = [];
    plan(wb).forEach((x) => items.push(Object.assign({ col: lastCol(wb, x.sheet) + (x.slot || 0) }, x)));
    fs.writeFileSync(full, embed(new Uint8Array(buf), items));
    console.log("embedded", items.length, "attachment(s) in", file);
  }

  await apply("ATS_RAID_Log.xlsx", (wb) => {
    const r = ATS.raid.parse(wb), row = (c, id) => r[c].find((x) => x.id === id).row;
    return [
      { type: "image", sheet: "Risks", row: row("risks", "R-001"), name: "environment-health.png", bytes: pic.env, cols: 3, rows: 5, slot: 0 },
      { type: "object", sheet: "Risks", row: row("risks", "R-001"), name: "Environment outage - SIT blocked.msg", bytes: MAIL.outage, slot: 3 },
      { type: "object", sheet: "Risks", row: row("risks", "R-003"), name: "Vendor build 3.2 - delivery date moving.msg", bytes: MAIL.vendor, slot: 0 },
      { type: "image", sheet: "Issues", row: row("issues", "I-001"), name: "environment-health.png", bytes: pic.env, cols: 3, rows: 5, slot: 0 },
      { type: "object", sheet: "Decisions", row: row("decisions", "DEC-003"), name: "Approved - move Cycle 2 start.msg", bytes: MAIL.decision, slot: 0 },
    ];
  });
  await apply("ATS_Weekly_Data.xlsx", (wb) => {
    const k = ATS.kpi.parse(wb), d = k.defects.find((x) => x.id === "SA-106"), t = k.tsrs[k.tsrs.length - 1], snap = k.snapshots[k.snapshots.length - 1];
    const out = [
      { type: "image", sheet: "Defect_Log", row: d.row, name: "SA-106-error.png", bytes: pic.err, cols: 3, rows: 5, slot: 0 },
      { type: "object", sheet: "Defect_Log", row: d.row, name: "SA-106 - submission timeout.msg", bytes: MAIL.defect, slot: 3 },
      { type: "object", sheet: "TSR_Log", row: t.row, name: "TSR response.msg", bytes: MAIL.tsr, slot: 0 },
      { type: "image", sheet: "Weekly_Snapshot", row: snap.row, name: "sit-burn-up.png", bytes: pic.burn, cols: 3, rows: 5, slot: 0 },
      { type: "image", sheet: "Lessons_Learned", row: k.lessons[0].row, name: "environment-health.png", bytes: pic.env, cols: 3, rows: 5, slot: 0 },
    ].concat(k.commentary.length ? [{ type: "object", sheet: "Commentary", row: k.commentary[0].row, name: "Steering group - decision.msg", bytes: MAIL.decision, slot: 0 }] : []);
    return out;
  });
  await apply("ATS_POAP.xlsx", (wb) => {
    const p = ATS.poap.parse(wb), bar = p.plan.find((b) => b.id === "A-018"), ms = p.milestones[0];
    return [
      { type: "image", sheet: "POAP_Plan", row: bar.row, name: "sit-burn-up.png", bytes: pic.burn, cols: 3, rows: 5, slot: 0 },
      { type: "object", sheet: "POAP_Milestones", row: ms.row, name: "Release 1 SIT start - entry criteria.msg", bytes: MAIL.plan, slot: 0 },
    ];
  });
})().catch((e) => { console.error(e); process.exit(1); });
