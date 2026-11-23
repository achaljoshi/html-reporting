/* Put pictures and files (e.g. Outlook .msg) INTO an .xlsx the way Excel stores them. Used by the tests and by tools/build_sample_attachments.js.
     picture:  floating image in xl/drawings (Insert > Pictures)
     object:   Insert > Object > Create from file (display as icon): sheet <oleObjects> + vmlDrawing + xl/embeddings/oleObjectN.bin (Package / Ole10Native)
   embed(xlsxBytes, [{type:'image'|'object', sheet, row, col, name, bytes, preview?}]) -> Uint8Array   (row/col are 1-based, like Excel)
   Node only. */
"use strict";
const path = require("path");
global.window = global.window || global;
const XLSX = global.XLSX || (global.XLSX = require(path.join(__dirname, "..", "assets", "lib", "xlsx.full.min.js")));

const enc = (s) => Buffer.from(s, "utf8");
const dec = (b) => Buffer.from(b).toString("utf8");
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

function readZip(bytes) {
  const z = XLSX.CFB.read(bytes, { type: "buffer" }), files = {};
  z.FullPaths.forEach((p, i) => { const n = p.replace(/^Root Entry\//, ""); if (!/\/$/.test(p) && n !== "\u0001Sh33tJ5" && z.FileIndex[i].content) files[n] = Buffer.from(z.FileIndex[i].content); });   // \x01Sh33tJ5 is SheetJS's own marker, not part of the package
  return files;
}
// Minimal zip writer (deflate via Node's zlib): SheetJS's built-in zip writer compresses sheet XML about 3x worse.
const zlib = require("zlib");
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
const crc32 = (b) => { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
function writeZip(files) {
  const names = Object.keys(files).sort((a, b) => (a === "[Content_Types].xml" ? -1 : b === "[Content_Types].xml" ? 1 : 0)), locals = [], central = [];
  let off = 0;
  names.forEach((name) => {
    const data = Buffer.from(files[name]), nm = Buffer.from(name, "utf8"), already = /\.(png|jpe?g|gif|zip)$/i.test(name), comp = already ? data : zlib.deflateRawSync(data, { level: 9 }), method = already ? 0 : 8, crc = crc32(data);
    const lh = Buffer.alloc(30); lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0x0800, 6); lh.writeUInt16LE(method, 8); lh.writeUInt16LE(0, 10); lh.writeUInt16LE(0x21, 12); lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(data.length, 22); lh.writeUInt16LE(nm.length, 26);
    const ch = Buffer.alloc(46); ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(0x0800, 8); ch.writeUInt16LE(method, 10); ch.writeUInt16LE(0x21, 14); ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(comp.length, 20); ch.writeUInt32LE(data.length, 24); ch.writeUInt16LE(nm.length, 28); ch.writeUInt32LE(off, 42);
    locals.push(lh, nm, comp); central.push(Buffer.concat([ch, nm])); off += 30 + nm.length + comp.length;
  });
  const cd = Buffer.concat(central), end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(names.length, 8); end.writeUInt16LE(names.length, 10); end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(off, 16);
  return new Uint8Array(Buffer.concat(locals.concat([cd, end])));
}

/** Ole10Native "Package" container exactly as Office writes it: size, flags, label, source path, temp path, payload */
function ole10Native(name, data) {
  const asciiz = (s) => Buffer.concat([Buffer.from(s, "latin1"), Buffer.from([0])]);
  const u32 = (n) => { const b = Buffer.alloc(4); b.writeUInt32LE(n); return b; };
  const src = asciiz("C:\\Users\\author\\Documents\\" + name), tmp = asciiz("C:\\Users\\author\\AppData\\Local\\Temp\\" + name);
  const body = Buffer.concat([Buffer.from([2, 0]), asciiz(name), src, u32(0x00030000), u32(tmp.length), tmp, u32(data.length), Buffer.from(data)]);
  return Buffer.concat([u32(body.length), body]);
}
function oleBin(name, data) {
  const c = XLSX.CFB.utils.cfb_new();
  XLSX.CFB.utils.cfb_add(c, "\u0001Ole10Native", ole10Native(name, data));
  XLSX.CFB.utils.cfb_add(c, "\u0001CompObj", Buffer.from([1, 0, 254, 255, 3, 10, 0, 0, 255, 255, 255, 255, 12, 0, 3, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70]));
  return Buffer.from(XLSX.CFB.write(c, { type: "buffer" }));
}

function addRel(files, relPath, type, target) {
  const xml = files[relPath] ? dec(files[relPath]) : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const used = [...xml.matchAll(/Id="([^"]+)"/g)].map((m) => m[1]);
  let n = 1; while (used.includes("rIdA" + n)) n++;
  const id = "rIdA" + n;
  files[relPath] = enc(xml.replace("</Relationships>", `<Relationship Id="${id}" Type="${REL}${type}" Target="${target}"/></Relationships>`));
  return id;
}
const relsOf = (sheetPath) => sheetPath.replace(/([^/]*)$/, "_rels/$1.rels");
const findRel = (files, relPath, type) => { const x = files[relPath] && dec(files[relPath]); const m = x && new RegExp(`<Relationship\\b[^>]*Type="[^"]*/${type}"[^>]*>`).exec(x); if (!m) return null; const t = /Target="([^"]+)"/.exec(m[0])[1], id = /Id="([^"]+)"/.exec(m[0])[1]; return { id, target: t }; };
const resolve = (dir, t) => { if (t.startsWith("/")) return t.slice(1); const out = []; (dir + t).split("/").forEach((p) => { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); }); return out.join("/"); };
const freePath = (files, make) => { let n = 1; while (files[make(n)]) n++; return make(n); };
const anchor = (col, row, w, h) => `<xdr:from><xdr:col>${col - 1}</xdr:col><xdr:colOff>38100</xdr:colOff><xdr:row>${row - 1}</xdr:row><xdr:rowOff>38100</xdr:rowOff></xdr:from><xdr:to><xdr:col>${col - 1 + w}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row - 1 + h}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>`;

function insertBefore(xml, tags, snippet) {
  let idx = -1;
  tags.forEach((t) => { const i = xml.indexOf("<" + t); if (i >= 0 && (idx < 0 || i < idx)) idx = i; });
  if (idx < 0) idx = xml.lastIndexOf("</worksheet>");
  return xml.slice(0, idx) + snippet + xml.slice(idx);
}
const AFTER_DRAWING = ["legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
const AFTER_LEGACY = ["legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"];
const AFTER_OLE = ["controls", "webPublishItems", "tableParts", "extLst"];

const VML_HEAD = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>`;
const VML_T75 = `<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f"><v:stroke joinstyle="miter"/><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/><o:lock v:ext="edit" aspectratio="t"/></v:shapetype>`;

const PNG_ICON = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==", "base64");

function embed(xlsx, items) {
  const files = readZip(xlsx), ct0 = dec(files["[Content_Types].xml"]);
  let ct = ct0;
  const addDefault = (ext, type) => { if (!new RegExp(`<Default[^>]*Extension="${ext}"`, "i").test(ct)) ct = ct.replace(/(<Types\b[^>]*>)/, `$1<Default Extension="${ext}" ContentType="${type}"/>`); };
  const addOverride = (p, type) => { if (!ct.includes(`PartName="${p}"`)) ct = ct.replace("</Types>", `<Override PartName="${p}" ContentType="${type}"/></Types>`); };
  const wb = dec(files["xl/workbook.xml"]), wbRels = dec(files["xl/_rels/workbook.xml.rels"]);
  const sheetPath = (name) => {
    const m = [...wb.matchAll(/<sheet\b([^>]*)\/?>/g)].find((x) => /name="([^"]*)"/.exec(x[1])[1].replace(/&amp;/g, "&") === name);
    if (!m) throw new Error("sheet not found: " + name);
    const rid = /r:id="([^"]+)"/.exec(m[1])[1], t = new RegExp(`<Relationship\\b[^>]*Id="${rid}"[^>]*>`).exec(wbRels)[0];
    return resolve("xl/", /Target="([^"]+)"/.exec(t)[1]);
  };
  const bySheet = {};
  items.forEach((it) => (bySheet[it.sheet] = bySheet[it.sheet] || []).push(it));
  let mediaN = 0;

  Object.entries(bySheet).forEach(([sheet, list]) => {
    const sp = sheetPath(sheet);
    let sx = dec(files[sp]);
    if (!/xmlns:r=/.test(sx)) sx = sx.replace("<worksheet ", `<worksheet xmlns:r="${NS_R}" `);
    const rp = relsOf(sp);
    // ---- pictures -> drawing
    const imgs = list.filter((x) => x.type === "image");
    if (imgs.length) {
      let dr = findRel(files, rp, "drawing"), dpath;
      if (dr) dpath = resolve(sp.replace(/[^/]*$/, ""), dr.target);
      else {
        dpath = freePath(files, (n) => `xl/drawings/drawing${n}.xml`);
        files[dpath] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}"></xdr:wsDr>`);
        const id = addRel(files, rp, "drawing", "../drawings/" + dpath.split("/").pop());
        sx = insertBefore(sx, AFTER_DRAWING, `<drawing r:id="${id}"/>`);
        addOverride("/" + dpath, "application/vnd.openxmlformats-officedocument.drawing+xml");
      }
      let dx = dec(files[dpath]);
      const drel = relsOf(dpath), ids = [...dx.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map((m) => +m[1]);
      let nextId = Math.max(1, ...ids) + 1;
      imgs.forEach((it) => {
        const ext = /^\x89PNG/.test(Buffer.from(it.bytes).toString("latin1", 0, 4)) ? "png" : "jpeg", mp = `xl/media/att${++mediaN}_${files["xl/media/image1.png"] ? "x" : ""}${Date.now() % 1e5}.${ext}`;
        files[mp] = Buffer.from(it.bytes); addDefault(ext, ext === "png" ? "image/png" : "image/jpeg");
        const rid = addRel(files, drel, "image", "../media/" + mp.split("/").pop());
        const w = it.cols || 4, h = it.rows || 8;
        dx = dx.replace("</xdr:wsDr>", `<xdr:twoCellAnchor editAs="oneCell">${anchor(it.col, it.row, w, h)}<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${nextId++}" name="${esc(it.name)}" descr="${esc(it.name)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w * 640000}" cy="${h * 190000}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`);
      });
      files[dpath] = enc(dx);
    }
    // ---- embedded objects -> oleObjects + vml
    const objs = list.filter((x) => x.type === "object");
    if (objs.length) {
      let lg = findRel(files, rp, "vmlDrawing"), vpath, vml;
      if (lg) { vpath = resolve(sp.replace(/[^/]*$/, ""), lg.target); vml = dec(files[vpath]); }
      else {
        vpath = freePath(files, (n) => `xl/drawings/vmlDrawing${n}.vml`); vml = VML_HEAD + VML_T75 + "</xml>";
        const id = addRel(files, rp, "vmlDrawing", "../drawings/" + vpath.split("/").pop());
        sx = insertBefore(sx, AFTER_LEGACY, `<legacyDrawing r:id="${id}"/>`);
        addDefault("vml", "application/vnd.openxmlformats-officedocument.vmlDrawing");
      }
      if (!vml.includes('id="_x0000_t75"')) vml = vml.replace(/(<o:shapelayout[\s\S]*?<\/o:shapelayout>)?/, (m) => m + VML_T75);
      const vrel = relsOf(vpath);
      let shapeId = Math.max(1024, ...[...vml.matchAll(/id="_x0000_s(\d+)"/g)].map((m) => +m[1])) + 1;
      let ole = "";
      objs.forEach((it) => {
        const bp = freePath(files, (n) => `xl/embeddings/oleObject${n}.bin`);
        files[bp] = oleBin(it.name, it.bytes); addDefault("bin", "application/vnd.openxmlformats-officedocument.oleObject");
        const pp = `xl/media/att${++mediaN}_${Date.now() % 1e5}.png`; files[pp] = Buffer.from(it.preview || PNG_ICON); addDefault("png", "image/png");
        const oid = addRel(files, rp, "oleObject", "../embeddings/" + bp.split("/").pop());
        const iid = addRel(files, rp, "image", "../media/" + pp.split("/").pop());
        const vid = addRel(files, vrel, "image", "../media/" + pp.split("/").pop());
        const sid = shapeId++;
        ole += `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><oleObject progId="Package" dvAspect="DVASPECT_ICON" shapeId="${sid}" r:id="${oid}"><objectPr defaultSize="0" r:id="${iid}"><anchor moveWithCells="1" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">${anchor(it.col, it.row, 1, 3)}</anchor></objectPr></oleObject></mc:Choice><mc:Fallback><oleObject progId="Package" dvAspect="DVASPECT_ICON" shapeId="${sid}" r:id="${oid}"/></mc:Fallback></mc:AlternateContent>`;
        vml = vml.replace("</xml>", `<v:shape id="_x0000_s${sid}" type="#_x0000_t75" style="position:absolute;margin-left:10pt;margin-top:10pt;width:48pt;height:48pt;z-index:${sid - 1000};visibility:visible" o:ole=""><v:imagedata o:relid="${vid}" o:title="${esc(it.name)}"/><x:ClientData ObjectType="Pict"><x:SizeWithCells/><x:Anchor>${it.col - 1}, 2, ${it.row - 1}, 2, ${it.col}, 2, ${it.row + 2}, 2</x:Anchor><x:CF>Pict</x:CF><x:AutoPict/><x:Row>${it.row - 1}</x:Row><x:Column>${it.col - 1}</x:Column></x:ClientData></v:shape></xml>`);
      });
      files[vpath] = enc(vml);
      if (/<oleObjects>/.test(sx)) sx = sx.replace("</oleObjects>", ole + "</oleObjects>");
      else sx = insertBefore(sx, AFTER_OLE, `<oleObjects>${ole}</oleObjects>`);
    }
    files[sp] = enc(sx);
  });
  files["[Content_Types].xml"] = enc(ct);
  return writeZip(files);
}

module.exports = { embed, oleBin, ole10Native, readZip, writeZip };
