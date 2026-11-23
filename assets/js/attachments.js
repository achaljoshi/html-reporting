/* ============================================================
   Attachments — pictures and Outlook messages (.msg) placed inside the Excel workbooks.

   Read straight out of the .xlsx package (no server, no upload):
     • floating pictures      xl/drawings/drawingN.xml  -> xl/media/*
     • embedded objects       sheet <oleObjects> -> xl/embeddings/oleObject*.bin  (Insert > Object > Create from file,
                              or an email dragged from Outlook) -> .msg / images / any file
     • pictures placed in a cell (Excel 365 "Place in cell")  via xl/richData
   Each attachment is tied to the sheet row it sits on, so it shows next to that note / defect / risk in the dashboard.

   ATS.att.extract(u8)            -> { items:[{sheet,row,col,name,kind,mime,size,b64,email?}], warnings:[] }
   ATS.att.link(kind, parsed, ex) -> sets parsed.attachments and record.att (ids) on every record whose `row` matches
   ATS.att.parseMsg(u8)           -> {subject, from, to[], cc[], date, bodyText, bodyHtml, attachments[]}
   Works in the browser and in Node (tests): needs only SheetJS (XLSX.CFB reads zip and OLE compound files).
   ============================================================ */
(function () {
  "use strict";
  const ATS = (window.ATS = window.ATS || {});
  const A = (ATS.att = ATS.att || {});
  const MAX_BYTES = 15 * 1024 * 1024;

  // ---------------------------------------------------------------- bytes helpers
  const u8 = (x) => (x instanceof Uint8Array ? x : x ? Uint8Array.from(x) : new Uint8Array(0));
  A.toB64 = function (bytes) {
    bytes = u8(bytes);
    if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(s);
  };
  A.fromB64 = function (b64) {
    if (typeof Buffer !== "undefined") return new Uint8Array(Buffer.from(b64, "base64"));
    const s = atob(b64), out = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
    return out;
  };
  const dec = (bytes, label) => { try { return new TextDecoder(label).decode(bytes); } catch (e) { return String.fromCharCode.apply(null, Array.from(bytes)); } };
  const utf8 = (b) => dec(b, "utf-8");
  const hash8 = (s) => { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return (h >>> 0).toString(16).padStart(8, "0"); };

  // ---------------------------------------------------------------- type sniffing
  A.sniff = function (b) {
    b = u8(b);
    const at = (o, ...v) => v.every((x, i) => b[o + i] === x);
    if (at(0, 0x89, 0x50, 0x4e, 0x47)) return { kind: "image", mime: "image/png", ext: "png" };
    if (at(0, 0xff, 0xd8, 0xff)) return { kind: "image", mime: "image/jpeg", ext: "jpg" };
    if (at(0, 0x47, 0x49, 0x46, 0x38)) return { kind: "image", mime: "image/gif", ext: "gif" };
    if (at(0, 0x42, 0x4d) && b.length > 30) return { kind: "image", mime: "image/bmp", ext: "bmp" };
    if (at(0, 0x52, 0x49, 0x46, 0x46) && at(8, 0x57, 0x45, 0x42, 0x50)) return { kind: "image", mime: "image/webp", ext: "webp" };
    if (at(0, 0x01, 0x00, 0x00, 0x00) && at(40, 0x20, 0x45, 0x4d, 0x46)) return { kind: "file", mime: "image/emf", ext: "emf" };
    if (at(0, 0xd7, 0xcd, 0xc6, 0x9a)) return { kind: "file", mime: "image/wmf", ext: "wmf" };
    if (at(0, 0x25, 0x50, 0x44, 0x46)) return { kind: "file", mime: "application/pdf", ext: "pdf" };
    if (at(0, 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1)) return { kind: "cfb", mime: "application/octet-stream", ext: "" };
    if (at(0, 0x50, 0x4b, 0x03, 0x04)) return { kind: "file", mime: "application/zip", ext: "zip" };
    return { kind: "file", mime: "application/octet-stream", ext: "" };
  };
  const MIME_BY_EXT = { png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", bmp: "image/bmp", webp: "image/webp", pdf: "application/pdf", msg: "application/vnd.ms-outlook", eml: "message/rfc822",
    txt: "text/plain", csv: "text/csv", xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation", zip: "application/zip", svg: "image/svg+xml", emf: "image/emf", wmf: "image/wmf" };
  // a MIME type that came from a file is untrusted text: only the plain type/subtype form is ever put into a page
  const safeMime = (m, fallback) => (/^[A-Za-z0-9][\w.+-]{0,60}\/[A-Za-z0-9][\w.+-]{0,100}$/.test(String(m || "")) ? String(m).toLowerCase() : fallback || "application/octet-stream");
  A.safeMime = safeMime;
  const extOf = (n) => { const m = /\.([A-Za-z0-9]{1,5})$/.exec(n || ""); return m ? m[1].toLowerCase() : ""; };

  // ---------------------------------------------------------------- Outlook .msg
  const readCfb = (bytes) => {
    const cfb = XLSX.CFB.read(u8(bytes), { type: "array" }), S = {};
    cfb.FullPaths.forEach((p, i) => { const e = cfb.FileIndex[i]; if (e && e.type === 2 && e.content) S[p.replace(/^Root Entry\//, "")] = u8(e.content); });
    return { cfb, S };
  };
  A.isMsg = (bytes) => { try { return A.sniff(bytes).kind === "cfb" && "__properties_version1.0" in readCfb(bytes).S; } catch (e) { return false; } };

  const filetime = (dv, o) => { const lo = dv.getUint32(o, true), hi = dv.getUint32(o + 4, true); const ms = (hi * 4294967296 + lo) / 10000 - 11644473600000; return ms > 0 && ms < 4e12 ? new Date(ms).toISOString() : null; };
  function props(bytes, header) {
    const out = {};
    if (!bytes || bytes.length < header + 16) return out;
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    for (let o = header; o + 16 <= bytes.length; o += 16) {
      const type = dv.getUint16(o, true), id = dv.getUint16(o + 2, true), key = id.toString(16).toUpperCase().padStart(4, "0");
      if (type === 0x0003) out[key] = dv.getInt32(o + 8, true);
      else if (type === 0x000b) out[key] = dv.getUint16(o + 8, true) !== 0;
      else if (type === 0x0040) out[key] = filetime(dv, o + 8);
    }
    return out;
  }

  // ---- compressed RTF (property 1009): LZFu decompression, then HTML recovered from RTF-encapsulated HTML
  const RTF_DICT = "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx";
  function lzfu(b) {
    if (!b || b.length < 16) return null;
    const dv = new DataView(b.buffer, b.byteOffset, b.byteLength), raw = dv.getUint32(4, true), magic = dec(b.subarray(8, 12), "latin1");
    if (magic === "MELA") return dec(b.subarray(16, 16 + raw), "windows-1252");
    if (magic !== "LZFu") return null;
    const dict = new Uint8Array(4096); for (let i = 0; i < RTF_DICT.length; i++) dict[i] = RTF_DICT.charCodeAt(i);
    let w = RTF_DICT.length, p = 16; const out = [];
    while (p < b.length && out.length < raw) {
      const ctl = b[p++];
      for (let bit = 0; bit < 8 && p < b.length && out.length < raw; bit++) {
        if (ctl & (1 << bit)) {
          const hi = b[p++], lo = b[p++]; if (lo === undefined) break;
          let off = (hi << 4) | (lo >> 4); const len = (lo & 15) + 2;
          if (off === (w & 4095)) return dec(Uint8Array.from(out), "windows-1252");           // end marker
          for (let k = 0; k < len; k++) { const c = dict[(off + k) & 4095]; out.push(c); dict[w & 4095] = c; w++; }
        } else { const c = b[p++]; out.push(c); dict[w & 4095] = c; w++; }
      }
    }
    return out.length ? dec(Uint8Array.from(out), "windows-1252") : null;
  }
  function rtfHtml(rtf) {
    if (!rtf || !/\\fromhtml/.test(rtf)) return "";
    let out = "", i = 0, st = { rtf: false, ign: false }; const stack = []; const n = rtf.length; let skipFallback = 0;
    const emit = (t) => { if (!st.rtf && !st.ign) out += t; };
    while (i < n) {
      const c = rtf[i];
      if (c === "{") { stack.push(st); st = { rtf: st.rtf, ign: st.ign }; i++; }
      else if (c === "}") { st = stack.pop() || st; i++; }
      else if (c === "\\") {
        const nx = rtf[i + 1];
        if (nx === "\\" || nx === "{" || nx === "}") { emit(nx); i += 2; }
        else if (nx === "'") { if (skipFallback > 0) skipFallback--; else emit(String.fromCharCode(parseInt(rtf.substr(i + 2, 2), 16))); i += 4; }
        else if (nx === "*") { i += 2; const m = /^\\([a-z]+)/.exec(rtf.slice(i, i + 20)); if (!(m && /^m?htmltag$/.test(m[1].replace(/\d+$/, "")))) { if (!(m && /^htmltag/.test(m[1]))) st.ign = true; } }
        else {
          const m = /^\\([a-z]+)(-?\d+)? ?/.exec(rtf.slice(i, i + 40));
          if (!m) { i += 2; continue; }
          const w = m[1], v = m[2] === undefined ? null : +m[2]; i += m[0].length;
          if (w === "htmlrtf") st.rtf = v !== 0;
          else if (/^(fonttbl|colortbl|stylesheet|info|header|footer|pict|object|generator)$/.test(w)) st.ign = true;
          else if (w === "u" && v !== null) { emit(String.fromCharCode(v < 0 ? v + 65536 : v)); skipFallback = 1; }
          else if (w === "tab") emit("\t");
          else if ((w === "par" || w === "line") && !st.rtf) emit("\n");
        }
      } else { if (c !== "\r" && c !== "\n") { if (skipFallback > 0) skipFallback--; else emit(c); } i++; }
    }
    return out.trim();
  }

  // Parse the message stored under `prefix` of the stream map S (prefix "" = the file itself; an attached email is nested under an attachment).
  function parseMessage(S, prefix, header) {
    const rd = (pre, tag) => {
      const w = S[pre + "__substg1.0_" + tag + "001F"], a = S[pre + "__substg1.0_" + tag + "001E"];
      if (w) return dec(w, "utf-16le").replace(/\0+$/, "");
      if (a) return dec(a, "windows-1252").replace(/\0+$/, "");
      return "";
    };
    const bin = (pre, tag) => S[pre + "__substg1.0_" + tag + "0102"] || null;
    const top = props(S[prefix + "__properties_version1.0"], header);
    const m = { subject: rd(prefix, "0037"), from: { name: rd(prefix, "0C1A"), email: rd(prefix, "5D01") || (/@/.test(rd(prefix, "0C1F")) ? rd(prefix, "0C1F") : "") || (/@/.test(rd(prefix, "0065")) ? rd(prefix, "0065") : "") },
      to: [], cc: [], bcc: [], date: top["0039"] || top["0E06"] || top["3007"] || null, bodyText: rd(prefix, "1000"), bodyHtml: "", attachments: [], headers: rd(prefix, "007D") };
    const ids = (re) => Array.from(new Set(Object.keys(S).filter((k) => k.startsWith(prefix)).map((k) => re.exec(k.slice(prefix.length))).filter(Boolean).map((r) => r[1]))).sort();
    ids(/^__recip_version1\.0_#([0-9A-Fa-f]{8})\//).forEach((id) => {
      const pre = prefix + "__recip_version1.0_#" + id + "/", pr = props(S[pre + "__properties_version1.0"], 8);
      const r = { name: rd(pre, "3001"), email: rd(pre, "39FE") || (/@/.test(rd(pre, "3003")) ? rd(pre, "3003") : "") };
      (pr["0C15"] === 2 ? m.cc : pr["0C15"] === 3 ? m.bcc : m.to).push(r);
    });
    if (!m.to.length && rd(prefix, "0E04")) rd(prefix, "0E04").split(";").map((x) => x.trim()).filter(Boolean).forEach((nm) => m.to.push({ name: nm, email: "" }));
    const hb = bin(prefix, "1013");
    if (hb) { const t = utf8(hb); m.bodyHtml = t.includes("�") ? dec(hb, "windows-1252") : t; }
    else m.bodyHtml = rd(prefix, "1013");
    if (!m.bodyHtml) { try { m.bodyHtml = rtfHtml(lzfu(bin(prefix, "1009"))); } catch (e) { /* plain text is shown instead */ } }
    if (!m.date && m.headers) { const dm = /^Date:\s*(.+)$/im.exec(m.headers); if (dm && !isNaN(Date.parse(dm[1]))) m.date = new Date(dm[1]).toISOString(); }
    ids(/^__attach_version1\.0_#([0-9A-Fa-f]{8})\//).forEach((id) => {
      const pre = prefix + "__attach_version1.0_#" + id + "/", pr = props(S[pre + "__properties_version1.0"], 8), data = bin(pre, "3701");
      const nested = !data && Object.keys(S).some((k) => k.startsWith(pre + "__substg1.0_3701000D/"));
      const name = rd(pre, "3707") || rd(pre, "3704") || rd(pre, "3001") || "attachment" + (rd(pre, "3703") || "");
      const sn = A.sniff(data || []);
      const att = { name, mime: safeMime(rd(pre, "370E"), sn.ext ? sn.mime : MIME_BY_EXT[extOf(name)] || "application/octet-stream"), size: data ? data.length : 0, cid: rd(pre, "3712"), embeddedMsg: nested, hidden: !!pr["7FFE"], b64: data ? A.toB64(data) : "" };
      if (nested) { att.email = parseMessage(S, pre + "__substg1.0_3701000D/", 24); att.mime = "application/vnd.ms-outlook"; att.name = (/\.msg$/i.test(name) ? name : (att.email.subject || name || "Email") + ".msg"); }
      m.attachments.push(att);
    });
    if (m.bodyHtml) m.attachments.forEach((a) => { if (a.cid && a.b64 && /^image\//.test(a.mime)) m.bodyHtml = m.bodyHtml.split("cid:" + a.cid).join("data:" + a.mime + ";base64," + a.b64); });
    return m;
  }
  A.parseMsg = function (bytes) { return parseMessage(readCfb(bytes).S, "", 32); };

  // RFC 5322 .eml built from a parsed message, so any mail client (Outlook, Apple Mail, Thunderbird) can open a message that was nested or whose .msg you do not want
  A.toEml = function (m) {
    const addr = (r) => (r.email ? (r.name ? '"' + r.name.replace(/"/g, "") + '" <' + r.email + ">" : r.email) : r.name || "");
    const enc = (s) => /^[\x20-\x7e]*$/.test(s || "") ? s || "" : "=?UTF-8?B?" + A.toB64(new TextEncoder().encode(s)) + "?=";
    const b64lines = (s) => (s.match(/.{1,76}/g) || []).join("\r\n");
    const L = [], bd = "=_ats_" + Date.now().toString(36);
    L.push("From: " + (addr(m.from) || "unknown@invalid"), "To: " + m.to.map(addr).join(", "));
    if (m.cc.length) L.push("Cc: " + m.cc.map(addr).join(", "));
    L.push("Subject: " + enc(m.subject), "Date: " + new Date(m.date || Date.now()).toUTCString(), "MIME-Version: 1.0", 'Content-Type: multipart/mixed; boundary="' + bd + '"', "", "--" + bd);
    if (m.bodyHtml) L.push('Content-Type: text/html; charset="utf-8"', "Content-Transfer-Encoding: base64", "", b64lines(A.toB64(new TextEncoder().encode(m.bodyHtml))));
    else L.push('Content-Type: text/plain; charset="utf-8"', "Content-Transfer-Encoding: base64", "", b64lines(A.toB64(new TextEncoder().encode(m.bodyText || ""))));
    m.attachments.filter((a) => a.b64).forEach((a) => L.push("--" + bd, 'Content-Type: ' + a.mime + '; name="' + enc(a.name) + '"', 'Content-Disposition: attachment; filename="' + enc(a.name) + '"', "Content-Transfer-Encoding: base64", "", b64lines(a.b64)));
    L.push("--" + bd + "--", "");
    return L.join("\r\n");
  };

  // ---------------------------------------------------------------- embedded OLE object (Package / .msg)
  const asciiz = (b, p) => { let e = p; while (e < b.length && b[e] !== 0) e++; return [dec(b.subarray(p, e), "windows-1252"), e + 1]; };
  function unpackOle(bytes) {
    let parsed;
    try { parsed = readCfb(bytes); } catch (e) { return { name: "object.bin", data: u8(bytes) }; }
    const S = parsed.S;
    if ("__properties_version1.0" in S) return { name: "", data: u8(bytes) };           // the object is the .msg itself (email dragged from Outlook)
    const key = Object.keys(S).find((k) => /Ole10Native$/.test(k));
    if (key) {
      const b = S[key], dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
      try {
        let [label, p] = asciiz(b, 6), src;
        [src, p] = asciiz(b, p);
        p += 4;
        const tl = dv.getUint32(p, true); p += 4 + tl;
        const size = dv.getUint32(p, true); p += 4;
        if (size > 0 && p + size <= b.length) return { name: label.split(/[\\/]/).pop() || src.split(/[\\/]/).pop(), data: b.subarray(p, p + size), src };
      } catch (e) { /* fall through to signature scan */ }
      for (let i = 0; i < Math.min(b.length - 8, 4096); i++) {            // fallback: first recognisable file signature
        const sn = A.sniff(b.subarray(i, i + 64));
        if (sn.kind !== "file" || sn.ext) { if (sn.kind !== "file" || sn.ext) return { name: "", data: b.subarray(i) }; }
      }
    }
    const pk = Object.keys(S).find((k) => /(^|\/)(Package|CONTENTS)$/.test(k));
    return { name: "", data: pk ? S[pk] : u8(bytes) };
  }

  // ---------------------------------------------------------------- xlsx package
  const attr = (s, n) => { const m = new RegExp("(?:^|\\s)" + n + '="([^"]*)"').exec(s); return m ? m[1] : ""; };
  const unxml = (s) => String(s || "").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
  function rels(files, partPath) {
    const dir = partPath.replace(/[^/]*$/, ""), file = partPath.slice(dir.length), r = files[dir + "_rels/" + file + ".rels"], out = {};
    if (!r) return out;
    const text = utf8(r), re = /<Relationship\b([^>]*)\/?>/g; let m;
    while ((m = re.exec(text))) { const a = m[1], id = attr(a, "Id"); let t = attr(a, "Target"); if (attr(a, "TargetMode") === "External") continue; out[id] = { type: attr(a, "Type").split("/").pop(), path: resolve(dir, t) }; }
    return out;
  }
  function resolve(dir, target) {
    if (target.startsWith("/")) return target.slice(1);
    const parts = (dir + target).split("/"), out = [];
    parts.forEach((p) => { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); });
    return out.join("/");
  }
  const fromRow = (xml) => { const m = /<(?:\w+:)?from>[\s\S]*?<(?:\w+:)?col>(\d+)<\/(?:\w+:)?col>[\s\S]*?<(?:\w+:)?row>(\d+)<\/(?:\w+:)?row>/.exec(xml); return m ? { col: +m[1] + 1, row: +m[2] + 1 } : null; };

  A.extract = function (input) {
    const out = { items: [], warnings: [] };
    let z;
    try { z = XLSX.CFB.read(u8(input), { type: "array" }); } catch (e) { return out; }
    const files = {};
    z.FullPaths.forEach((p, i) => { if (!/\/$/.test(p) && z.FileIndex[i].content) files[p.replace(/^Root Entry\//, "")] = u8(z.FileIndex[i].content); });
    if (!Object.keys(files).some((p) => /^xl\/(media|embeddings|drawings|richData)\//.test(p))) return out;       // fast exit: nothing attached
    const wbXml = files["xl/workbook.xml"] && utf8(files["xl/workbook.xml"]); if (!wbXml) return out;
    const wbRels = rels(files, "xl/workbook.xml");
    const sheets = []; { const re = /<sheet\b([^>]*)\/?>/g; let m; while ((m = re.exec(wbXml))) { const rid = attr(m[1], "r:id"); if (wbRels[rid]) sheets.push({ name: unxml(attr(m[1], "name")), path: wbRels[rid].path }); } }

    const push = (sheet, row, col, name, bytes, hint) => {
      if (!bytes || !bytes.length) return;
      if (bytes.length > MAX_BYTES) { out.warnings.push(`${sheet} row ${row}: "${name}" is larger than ${MAX_BYTES / 1048576} MB and was skipped.`); return; }
      const sn = A.sniff(bytes);
      let kind = sn.kind, mime = sn.mime, nm = String(name || "").replace(/[\u0000-\u001f]+/g, " ").replace(/\s+/g, " ").trim();
      if (kind === "cfb" && A.isMsg(bytes)) kind = "msg";
      else if (kind === "cfb") kind = "file";
      if (kind === "msg") { mime = "application/vnd.ms-outlook"; }
      else if (kind === "file" && !sn.ext && extOf(nm)) { mime = MIME_BY_EXT[extOf(nm)] || mime; if (/^image\/(png|jpe?g|gif|bmp|webp)$/.test(mime)) kind = "image"; }
      const item = { sheet, row, col, name: nm, kind, mime, size: bytes.length, b64: A.toB64(bytes) };
      if (kind === "msg") {
        try {
          const em = A.parseMsg(bytes);
          item.email = em;
          if (!item.name || /\.bin$|^object/i.test(item.name)) item.name = (em.subject || "Email").replace(/[\\/:*?"<>|]+/g, " ").trim().slice(0, 80) + ".msg";
        } catch (e) { out.warnings.push(`${sheet} row ${row}: an Outlook message could not be read (${e.message}); it is kept as a downloadable file.`); item.kind = "file"; }
      }
      if (!item.name) item.name = (hint || "attachment") + (sn.ext ? "." + sn.ext : "");
      if (!extOf(item.name) && sn.ext && kind !== "msg") item.name += "." + sn.ext;
      if (kind === "file" && /^image\/(emf|wmf)$/.test(mime)) item.previewOnly = true;
      out.items.push(item);
    };

    sheets.forEach((sh) => {
      const xml = files[sh.path] && utf8(files[sh.path]); if (!xml) return;
      const sr = rels(files, sh.path);
      // 1) floating pictures in the sheet's drawing
      Object.keys(sr).filter((k) => sr[k].type === "drawing").forEach((k) => {
        const dp = sr[k].path, dxml = files[dp] && utf8(files[dp]); if (!dxml) return;
        const dr = rels(files, dp), re = /<(?:\w+:)?(?:twoCellAnchor|oneCellAnchor)\b[\s\S]*?<\/(?:\w+:)?(?:twoCellAnchor|oneCellAnchor)>/g; let m;
        while ((m = re.exec(dxml))) {
          const blk = m[0], pos = fromRow(blk), blip = /<(?:\w+:)?blip\b[^>]*?r:embed="([^"]+)"/.exec(blk);
          if (!pos || !blip || !dr[blip[1]]) continue;
          const nv = /<(?:\w+:)?cNvPr\b([^>]*)\/?>/.exec(blk), nm = nv ? unxml(attr(nv[1], "descr") || attr(nv[1], "name")) : "";
          const media = files[dr[blip[1]].path];
          push(sh.name, pos.row, pos.col, /\.[A-Za-z0-9]{2,4}$/.test(nm) ? nm : "", media, nm || "Picture");
        }
      });
      // 2) embedded objects
      const vmlKey = Object.keys(sr).find((k) => sr[k].type === "vmlDrawing"), vml = vmlKey && files[sr[vmlKey].path] ? utf8(files[sr[vmlKey].path]) : "";
      const seen = new Set(), ore = /<oleObject\b([^>]*?)(?:\/>|>([\s\S]*?)<\/oleObject>)/g; let om;
      while ((om = ore.exec(xml))) {
        const a = om[1], rid = attr(a, "r:id"), shapeId = attr(a, "shapeId");
        let pos = om[2] ? fromRow(om[2]) : null;
        if (!pos && shapeId && vml) {                                                    // legacy anchor in the VML drawing
          const vs = new RegExp('<v:shape\\b[^>]*?id="_x0000_s' + shapeId + '"[\\s\\S]*?</v:shape>').exec(vml);
          if (vs) { const r = /<x:Row>(\d+)<\/x:Row>/.exec(vs[0]), c = /<x:Column>(\d+)<\/x:Column>/.exec(vs[0]), an = /<x:Anchor>\s*([^<]+)</.exec(vs[0]); if (r && c) pos = { row: +r[1] + 1, col: +c[1] + 1 }; else if (an) { const n = an[1].split(",").map((x) => +x.trim()); pos = { row: n[2] + 1, col: n[0] + 1 }; } }
        }
        if (seen.has(rid) && !pos) continue;
        if (!sr[rid] || !pos) { if (!seen.has(rid) && sr[rid]) out.warnings.push(`${sh.name}: an embedded object has no readable position; it was skipped.`); continue; }
        if (seen.has(rid)) continue;
        seen.add(rid);
        const raw = files[sr[rid].path]; if (!raw) continue;
        if (sr[rid].type === "package") { push(sh.name, pos.row, pos.col, sr[rid].path.split("/").pop(), raw, "Embedded file"); continue; }
        const o = unpackOle(raw);
        push(sh.name, pos.row, pos.col, o.name, o.data, attr(a, "progId") === "Package" ? "Embedded file" : attr(a, "progId") || "Embedded object");
      }
      // 3) pictures placed in cells (Excel 365): <c r="B5" vm="1"> -> metadata -> richData -> media
      const cellVm = []; { const cre = /<c\b([^>]*\bvm="(\d+)"[^>]*)\/?>/g; let cm; while ((cm = cre.exec(xml))) { const ref = /^([A-Z]+)(\d+)$/.exec(attr(cm[1], "r")); if (ref) cellVm.push({ vm: +cm[2], row: +ref[2], col: ref[1].split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) }); } }
      if (cellVm.length && files["xl/richData/rdrichvalue.xml"] && files["xl/richData/richValueRel.xml"] && files["xl/metadata.xml"]) {
        try {
          const meta = utf8(files["xl/metadata.xml"]), rv = utf8(files["xl/richData/rdrichvalue.xml"]), rel = utf8(files["xl/richData/richValueRel.xml"]), rr = rels(files, "xl/richData/richValueRel.xml");
          const vmToRv = []; { const vmeta = /<valueMetadata\b[\s\S]*?<\/valueMetadata>/.exec(meta); const bre = /<bk>\s*<rc\b[^>]*\bv="(\d+)"/g; let b; while (vmeta && (b = bre.exec(vmeta[0]))) vmToRv.push(+b[1]); }
          const rvb = []; { const fm = /<futureMetadata name="XLRICHVALUE"[\s\S]*?<\/futureMetadata>/.exec(meta); const re2 = /<xlrd:rvb i="(\d+)"/g; let b; while (fm && (b = re2.exec(fm[0]))) rvb.push(+b[1]); }
          const rvs = []; { const re3 = /<rv\b[^>]*>([\s\S]*?)<\/rv>/g; let b; while ((b = re3.exec(rv))) { const vs = [...b[1].matchAll(/<v>([^<]*)<\/v>/g)].map((x) => x[1]); rvs.push(+vs[0]); } }
          const rels2 = [...rel.matchAll(/<rel\b[^>]*r:id="([^"]+)"/g)].map((x) => x[1]);
          cellVm.forEach((c) => { const k = vmToRv[c.vm - 1], idx = k != null ? rvb[k] : null, relIdx = idx != null ? rvs[idx] : null, rid = relIdx != null ? rels2[relIdx] : null; if (rid && rr[rid]) push(sh.name, c.row, c.col, "", files[rr[rid].path], "Picture in cell"); });
        } catch (e) { out.warnings.push(`${sh.name}: pictures placed in cells could not be read (${e.message}).`); }
      }
    });
    return out;
  };

  // ---------------------------------------------------------------- linking to parsed records
  const COLLECTIONS = {
    kpi: { snapshots: "Weekly_Snapshot", topicProgress: "Topic_Progress", defects: "Defect_Log", tsrs: "TSR_Log", milestones: "Milestones", readiness: "Readiness", commentary: "Commentary", resourcing: "Resourcing", demand: "Demand_Forecast", lessons: "Lessons_Learned" },
    raid: { risks: "Risks", issues: "Issues", assumptions: "Assumptions", deps: "Dependencies", decisions: "Decisions" },
    poap: { plan: "POAP_Plan", milestones: "POAP_Milestones", p2p: "P2P_Matrix" },
  };
  // human label of the record an attachment is attached to (shown in the viewer and the Evidence library)
  const T = (x, n) => { x = String(x == null ? "" : x).replace(/\s+/g, " ").trim(); return x.length > n ? x.slice(0, n - 1) + "…" : x; };
  const LABELS = {
    Defect_Log: (r) => `Defect ${r.id} — ${T(r.summary, 60)}`, TSR_Log: (r) => `TSR ${r.ref}`, Milestones: (r) => `Milestone — ${T(r.name, 60)}`, Lessons_Learned: (r) => `Lesson — ${T(r.lesson, 60)}`,
    Commentary: (r) => `Commentary — ${r.kpi} (w/e ${r.weekEnding})`, Weekly_Snapshot: (r) => `Weekly snapshot, w/e ${r.weekEnding}`, Topic_Progress: (r) => `${r.topic} — w/e ${r.weekEnding}`,
    Readiness: (r) => `Readiness — ${T(r.topic, 50)} (w/e ${r.weekEnding})`, Resourcing: (r) => `Resourcing — ${T(r.role, 50)}`, Demand_Forecast: (r) => `Demand forecast — ${r.month}`,
    Risks: (r) => `Risk ${r.id} — ${T(r.title, 60)}`, Issues: (r) => `Issue ${r.id} — ${T(r.title, 60)}`, Assumptions: (r) => `Assumption ${r.id} — ${T(r.desc, 60)}`,
    Dependencies: (r) => `Dependency ${r.id} — ${T(r.desc, 60)}`, Decisions: (r) => `Decision ${r.id} — ${T(r.title, 60)}`,
    POAP_Plan: (r) => `Plan item ${r.id} — ${T(r.item, 60)}`, POAP_Milestones: (r) => `Milestone — ${T(r.name, 60)}`, P2P_Matrix: (r) => `Flow ${r.id || ""} — ${T(r.path || r.subProcess, 60)}`,
  };
  A.link = function (kind, parsed, ex, fileName) {
    if (!parsed) return parsed;
    const items = ((ex && ex.items) || []).map((it, i) => Object.assign({ id: hash8([fileName, it.sheet, it.row, it.col, i, it.name].join("|")) }, it));
    parsed.attachments = items;
    parsed.attachmentWarnings = (ex && ex.warnings) || [];
    if (!items.length) return parsed;
    const bySheetRow = {};
    items.forEach((it) => { ((bySheetRow[it.sheet] = bySheetRow[it.sheet] || {})[it.row] = (bySheetRow[it.sheet][it.row] || [])).push(it); });
    Object.entries(COLLECTIONS[kind] || {}).forEach(([prop, sheet]) => {
      (parsed[prop] || []).forEach((rec) => { const rows = bySheetRow[sheet]; const hit = rows && rec && rec.row != null ? rows[rec.row] : null; if (hit) { rec.att = hit.map((x) => x.id); hit.forEach((x) => { try { x.owner = LABELS[sheet] ? LABELS[sheet](rec) : sheet; } catch (e) { x.owner = sheet; } }); } });
    });
    items.forEach((it) => { it.linked = !!(COLLECTIONS[kind] && Object.values(COLLECTIONS[kind]).includes(it.sheet) && ((parsed[Object.keys(COLLECTIONS[kind]).find((p) => COLLECTIONS[kind][p] === it.sheet)] || []).some((r) => r.row === it.row))); });
    return parsed;
  };
  A.COLLECTIONS = COLLECTIONS;
})();
