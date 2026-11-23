/* Builds a minimal Outlook .msg (OLE compound file with MAPI property streams) from plain fields.
   Used for tests and the fictional demo emails; the reader in assets/js/attachments.js was verified against real Outlook files.
   makeMsg({subject, from:{name,email}, to:[{name,email}], cc:[], date:'2026-09-24T09:30:00Z', body, html, attachments:[{name,mime,bytes}]}) -> Buffer */
"use strict";
const path = require("path");
global.window = global.window || global;
const XLSX = global.XLSX || (global.XLSX = require(path.join(__dirname, "..", "assets", "lib", "xlsx.full.min.js")));
const U = XLSX.CFB.utils;

const utf16 = (s) => Buffer.from(String(s) + "\0", "utf16le");
function entry(type, id, value, flags) {
  const b = Buffer.alloc(16); b.writeUInt16LE(type, 0); b.writeUInt16LE(id, 2); b.writeUInt32LE(flags == null ? 6 : flags, 4);
  if (Buffer.isBuffer(value)) value.copy(b, 8); else if (typeof value === "number") b.writeInt32LE(value, 8);
  return b;
}
const intEntry = (id, v) => entry(0x0003, id, v);
const boolEntry = (id, v) => { const b = entry(0x000b, id, null); b.writeUInt16LE(v ? 1 : 0, 8); return b; };
function timeEntry(id, iso) {
  const ft = BigInt(new Date(iso).getTime()) * 10000n + 116444736000000000n, b = Buffer.alloc(8); b.writeBigUInt64LE(ft);
  return entry(0x0040, id, b);
}
const strEntry = (id, s, size) => { const v = Buffer.alloc(8); v.writeUInt32LE(size, 0); return entry(0x001f, id, v); };
const binEntry = (id, size) => { const v = Buffer.alloc(8); v.writeUInt32LE(size, 0); return entry(0x0102, id, v); };
const hex4 = (n) => n.toString(16).toUpperCase().padStart(4, "0");

function makeMsg(m) {
  const c = U.cfb_new(), add = (p, data) => U.cfb_add(c, "/" + p, Buffer.isBuffer(data) ? data : Buffer.from(data));
  const to = m.to || [], cc = m.cc || [], att = m.attachments || [], rcpts = to.map((r) => [r, 1]).concat(cc.map((r) => [r, 2]));
  const top = [];
  const str = (id, s) => { if (!s) return; const d = utf16(s); add(`__substg1.0_${hex4(id)}001F`, d); top.push(strEntry(id, s, d.length)); };
  add("__substg1.0_001A001F", utf16("IPM.Note")); top.push(strEntry(0x001a, "", 18));
  str(0x0037, m.subject); str(0x0C1A, m.from && m.from.name); str(0x5D01, m.from && m.from.email); str(0x0065, m.from && m.from.email);
  str(0x0E04, to.map((r) => r.name || r.email).join("; ")); str(0x0E03, cc.map((r) => r.name || r.email).join("; "));
  str(0x1000, m.body || "");
  if (m.html) { const h = Buffer.from(m.html, "utf8"); add("__substg1.0_10130102", h); top.push(binEntry(0x1013, h.length)); }
  if (m.date) { top.push(timeEntry(0x0039, m.date)); top.push(timeEntry(0x0E06, m.date)); }
  top.push(intEntry(0x0E07, 1));
  const hdr = Buffer.alloc(32); hdr.writeUInt32LE(rcpts.length, 8); hdr.writeUInt32LE(att.length, 12); hdr.writeUInt32LE(rcpts.length, 16); hdr.writeUInt32LE(att.length, 20);
  add("__properties_version1.0", Buffer.concat([hdr].concat(top)));
  rcpts.forEach(([r, type], i) => {
    const pre = `__recip_version1.0_#${String(i).padStart(8, "0")}/`, ps = [intEntry(0x0C15, type), intEntry(0x3000, i)];
    const name = utf16(r.name || r.email), mail = utf16(r.email || "");
    add(pre + "__substg1.0_3001001F", name); ps.push(strEntry(0x3001, "", name.length));
    add(pre + "__substg1.0_39FE001F", mail); ps.push(strEntry(0x39FE, "", mail.length));
    add(pre + "__properties_version1.0", Buffer.concat([Buffer.alloc(8)].concat(ps)));
  });
  att.forEach((a, i) => {
    const pre = `__attach_version1.0_#${String(i).padStart(8, "0")}/`, ps = [intEntry(0x3705, 1), intEntry(0x0FF4, 0)];
    const nm = utf16(a.name), data = Buffer.from(a.bytes), mime = utf16(a.mime || "application/octet-stream");
    add(pre + "__substg1.0_3707001F", nm); ps.push(strEntry(0x3707, "", nm.length));
    add(pre + "__substg1.0_3704001F", nm); ps.push(strEntry(0x3704, "", nm.length));
    add(pre + "__substg1.0_370E001F", mime); ps.push(strEntry(0x370E, "", mime.length));
    add(pre + "__substg1.0_37010102", data); ps.push(binEntry(0x3701, data.length));
    add(pre + "__properties_version1.0", Buffer.concat([Buffer.alloc(8)].concat(ps)));
  });
  return Buffer.from(XLSX.CFB.write(c, { type: "buffer" }));
}
module.exports = { makeMsg };
