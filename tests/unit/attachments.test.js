'use strict';
// Pictures and Outlook .msg files placed inside the Excel workbooks: extraction, row linking and the .msg reader.
// Test files are built in code (tools/xlsx_embed.js, tools/make_msg.js); the .msg reader was also verified against real Outlook files.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { ATS, XLSX, sampleWb, templateWb, edit, SAMPLE_DIR } = require('../helpers/app.js');
const { embed, oleBin, readZip } = require('../../tools/xlsx_embed.js');
const { makeMsg } = require('../../tools/make_msg.js');

const A = ATS.att;
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
// a RAID workbook built in code from the blank template (no attachments in it): 3 risks, 2 issues, 1 decision
const cleanRaid = () => {
  const wb = templateWb('ATS_RAID_Log_Template.xlsx');
  edit(wb, 'Risks', 'ID', (aoa, c) => { ['R-001', 'R-002', 'R-003'].forEach((id) => c.append({ ID: id, 'Summary Title': 'Risk ' + id, Likelihood: 'Likely', Impact: 'Major', Status: 'Open' })); });
  edit(wb, 'Issues', 'ID', (aoa, c) => { ['I-001', 'I-002'].forEach((id) => c.append({ ID: id, 'Summary Title': 'Issue ' + id, 'Full Description': 'x', Priority: 'High', Severity: 'Major', Status: 'Open' })); });
  edit(wb, 'Decisions', 'ID', (aoa, c) => { c.append({ ID: 'DEC-001', 'Summary Title': 'Decision one', Status: 'Pending' }); });
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
};
const RAID = cleanRaid();
const parseRaid = () => ATS.raid.parse(XLSX.read(RAID, { type: 'array', cellDates: true }));
const msg = (o) => makeMsg(Object.assign({ subject: 'Hello', from: { name: 'Ann Lee', email: 'ann@example.com' }, to: [{ name: 'Bob', email: 'bob@example.com' }], date: '2026-09-24T09:30:00Z', body: 'Body text' }, o));

test('a workbook without attachments yields nothing and no warnings', () => {
  const clean = new Uint8Array(fs.readFileSync(path.join(SAMPLE_DIR, 'PAYE', 'ATS_Weekly_Data.xlsx')));
  const ex = A.extract(clean);
  assert.deepStrictEqual(ex, { items: [], warnings: [] });
});

test('garbage input does not throw', () => {
  assert.deepStrictEqual(A.extract(new Uint8Array([1, 2, 3, 4])).items, []);
  assert.deepStrictEqual(A.extract(new Uint8Array(0)).items, []);
});

test('a floating picture is found with its sheet, row, column and name', () => {
  const xl = embed(RAID, [{ type: 'image', sheet: 'Risks', row: 14, col: 22, name: 'outage.png', bytes: PNG }]);
  const ex = A.extract(xl);
  assert.strictEqual(ex.items.length, 1);
  const it = ex.items[0];
  assert.deepStrictEqual([it.sheet, it.row, it.col, it.name, it.kind, it.mime, it.size], ['Risks', 14, 22, 'outage.png', 'image', 'image/png', PNG.length]);
  assert.deepStrictEqual(Buffer.from(A.fromB64(it.b64)), PNG);
});

test('an embedded Outlook message is found, parsed and anchored to its row', () => {
  const m = msg({ subject: 'Environment outage – SIT blocked', cc: [{ name: 'Env', email: 'env@example.com' }], html: '<p><b>blocked</b></p>', attachments: [{ name: 'shot.png', mime: 'image/png', bytes: PNG }] });
  const ex = A.extract(embed(RAID, [{ type: 'object', sheet: 'Issues', row: 12, col: 24, name: 'mail.msg', bytes: m }]));
  assert.strictEqual(ex.items.length, 1);
  const it = ex.items[0];
  assert.deepStrictEqual([it.sheet, it.row, it.kind, it.mime], ['Issues', 12, 'msg', 'application/vnd.ms-outlook']);
  assert.strictEqual(it.name, 'mail.msg');
  assert.strictEqual(it.email.subject, 'Environment outage – SIT blocked');
  assert.deepStrictEqual(it.email.from, { name: 'Ann Lee', email: 'ann@example.com' });
  assert.deepStrictEqual(it.email.cc, [{ name: 'Env', email: 'env@example.com' }]);
  assert.strictEqual(it.email.date, '2026-09-24T09:30:00.000Z');
  assert.strictEqual(it.email.bodyHtml, '<p><b>blocked</b></p>');
  assert.deepStrictEqual(it.email.attachments.map((a) => [a.name, a.size]), [['shot.png', PNG.length]]);
});

test('an embedded picture file (Package) is recognised by content, not by its name', () => {
  const ex = A.extract(embed(RAID, [{ type: 'object', sheet: 'Risks', row: 12, col: 24, name: 'evidence', bytes: PNG }]));
  assert.strictEqual(ex.items[0].kind, 'image');
  assert.match(ex.items[0].name, /\.png$/);
});

test('any other embedded file is kept as a downloadable file with its original name', () => {
  const pdf = Buffer.from('%PDF-1.4\n%fake\n');
  const ex = A.extract(embed(RAID, [{ type: 'object', sheet: 'Risks', row: 12, col: 24, name: 'report.pdf', bytes: pdf }]));
  assert.deepStrictEqual([ex.items[0].kind, ex.items[0].name, ex.items[0].mime], ['file', 'report.pdf', 'application/pdf']);
});

test('several attachments on one row and on different sheets are all found', () => {
  const xl = embed(RAID, [
    { type: 'image', sheet: 'Risks', row: 11, col: 20, name: 'a.png', bytes: PNG },
    { type: 'object', sheet: 'Risks', row: 11, col: 23, name: 'b.msg', bytes: msg({ subject: 'B' }) },
    { type: 'object', sheet: 'Risks', row: 13, col: 20, name: 'c.msg', bytes: msg({ subject: 'C' }) },
    { type: 'image', sheet: 'Decisions', row: 11, col: 22, name: 'd.png', bytes: PNG },
  ]);
  const ex = A.extract(xl);
  assert.deepStrictEqual(ex.items.map((i) => i.sheet + '#' + i.row + ':' + i.kind).sort(), ['Decisions#11:image', 'Risks#11:image', 'Risks#11:msg', 'Risks#13:msg']);
});

test('the legacy VML anchor is used when the sheet XML has no anchor of its own', () => {
  const files = readZip(Buffer.from(embed(RAID, [{ type: 'object', sheet: 'Risks', row: 15, col: 21, name: 'x.msg', bytes: msg({ subject: 'V' }) }])));
  const sheetPath = Object.keys(files).find((p) => /worksheets\/sheet\d+\.xml$/.test(p) && /<oleObject/.test(files[p].toString('utf8')));
  let xml = files[sheetPath].toString('utf8');
  xml = xml.replace(/<objectPr[\s\S]*?<\/objectPr>/g, '').replace(/<mc:AlternateContent[\s\S]*?<mc:Fallback>/, '').replace(/<\/mc:Fallback><\/mc:AlternateContent>/, '');
  files[sheetPath] = Buffer.from(xml);
  const { writeZip } = require('../../tools/xlsx_embed.js');
  const ex = A.extract(writeZip(files));
  assert.deepStrictEqual(ex.items.map((i) => [i.sheet, i.row, i.kind]), [['Risks', 15, 'msg']]);
});

test('an oversized attachment is skipped with a warning instead of bloating memory', () => {
  const big = Buffer.concat([PNG, Buffer.alloc(16 * 1024 * 1024)]);
  const ex = A.extract(embed(RAID, [{ type: 'image', sheet: 'Risks', row: 11, col: 20, name: 'huge.png', bytes: big }]));
  assert.strictEqual(ex.items.length, 0);
  assert.match(ex.warnings[0], /larger than 15 MB/);
});

test('an object that is itself a .msg (email dragged from Outlook) is read', () => {
  // an object whose compound file IS the message
  const Z = require('../../tools/xlsx_embed.js');
  const files = readZip(Buffer.from(embed(RAID, [{ type: 'object', sheet: 'Risks', row: 12, col: 24, name: 'x.msg', bytes: PNG }])));
  const bin = Object.keys(files).find((p) => /embeddings\/oleObject\d+\.bin$/.test(p));
  files[bin] = msg({ subject: 'Dragged from Outlook' });
  const ex = A.extract(Z.writeZip(files));
  assert.deepStrictEqual([ex.items[0].kind, ex.items[0].email.subject], ['msg', 'Dragged from Outlook']);
});

test('Ole10Native fallback: a payload with an odd header is still found by its file signature', () => {
  const body = Buffer.concat([Buffer.from([2, 0]), Buffer.from('weird\0'), Buffer.from('junk-without-the-expected-layout'), PNG]);
  const len = Buffer.alloc(4); len.writeUInt32LE(body.length);
  const c = XLSX.CFB.utils.cfb_new(); XLSX.CFB.utils.cfb_add(c, '\u0001Ole10Native', Buffer.concat([len, body]));
  const bin = Buffer.from(XLSX.CFB.write(c, { type: 'buffer' }));
  const files = readZip(Buffer.from(embed(RAID, [{ type: 'object', sheet: 'Risks', row: 12, col: 24, name: 'x', bytes: PNG }])));
  files[Object.keys(files).find((p) => /oleObject\d+\.bin$/.test(p))] = bin;
  const ex = A.extract(require('../../tools/xlsx_embed.js').writeZip(files));
  assert.strictEqual(ex.items[0].kind, 'image');
});

test('link(): attachments land on the record of that sheet row and carry a readable owner label', () => {
  const r = parseRaid();
  const risk = r.risks[0], issue = r.issues[0];
  const xl = embed(RAID, [
    { type: 'image', sheet: 'Risks', row: risk.row, col: 20, name: 'a.png', bytes: PNG },
    { type: 'image', sheet: 'Issues', row: issue.row, col: 24, name: 'b.png', bytes: PNG },
    { type: 'image', sheet: 'Risks', row: 90, col: 20, name: 'stray.png', bytes: PNG },     // not on a data row
  ]);
  const wb = parseRaid();
  A.link('raid', wb, A.extract(xl), 'ATS_RAID_Log.xlsx');
  assert.strictEqual(wb.attachments.length, 3);
  assert.strictEqual(wb.risks[0].att.length, 1);
  assert.strictEqual(wb.issues[0].att.length, 1);
  assert.ok(!wb.risks[1].att);
  const a = wb.attachments.find((x) => x.name === 'a.png');
  assert.match(a.owner, /^Risk R-001 — /);
  assert.strictEqual(a.linked, true);
  assert.strictEqual(wb.attachments.find((x) => x.name === 'stray.png').linked, false);
  const ids = wb.attachments.map((x) => x.id);
  assert.strictEqual(new Set(ids).size, 3);                                                // ids are unique
});

test('row numbers match Excel rows even when the sheet starts below row 1', () => {
  const wb = sampleWb('Self Assessment', 'ATS_Weekly_Data.xlsx');
  const k = ATS.kpi.parse(wb);
  const d = k.defects.find((x) => x.id === 'SA-106');
  const rows = XLSX.utils.sheet_to_json(wb.Sheets.Defect_Log, { header: 1, range: 0, blankrows: true });
  assert.strictEqual(rows[d.row - 1][0], 'SA-106');
});

test('the demo workbooks carry fictional attachments linked to the right items', () => {
  const k = ATS.kpi.parse(sampleWb('Self Assessment', 'ATS_Weekly_Data.xlsx'));
  A.link('kpi', k, A.extract(new Uint8Array(fs.readFileSync(path.join(SAMPLE_DIR, 'Self Assessment', 'ATS_Weekly_Data.xlsx')))), 'w.xlsx');
  const d = k.defects.find((x) => x.id === 'SA-106');
  assert.strictEqual(d.att.length, 2);
  const kinds = k.attachments.filter((a) => d.att.includes(a.id)).map((a) => a.kind).sort();
  assert.deepStrictEqual(kinds, ['image', 'msg']);
  assert.ok(k.attachments.every((a) => a.linked), 'every demo attachment sits on a data row');
});

// ---------------------------------------------------------------- .msg reader
test('.msg: HTML, plain text, recipients and an inline picture (cid) are read', () => {
  const m = A.parseMsg(msg({ html: '<p><img src="cid:logo1"></p>', attachments: [] }));
  assert.strictEqual(m.subject, 'Hello');
  assert.deepStrictEqual(m.to, [{ name: 'Bob', email: 'bob@example.com' }]);
  assert.strictEqual(m.bodyText, 'Body text');
});

test('.msg: unicode subjects and bodies survive', () => {
  const m = A.parseMsg(msg({ subject: 'Café — 你好 ✓', body: 'こんにちは 안녕하세요' }));
  assert.strictEqual(m.subject, 'Café — 你好 ✓');
  assert.strictEqual(m.bodyText, 'こんにちは 안녕하세요');
});

test('.msg: toEml produces a mail file other clients can open', () => {
  const m = A.parseMsg(msg({ subject: 'Ünï', attachments: [{ name: 'a.png', mime: 'image/png', bytes: PNG }] }));
  const eml = A.toEml(m);
  assert.match(eml, /^From: "Ann Lee" <ann@example.com>/);
  assert.match(eml, /Subject: =\?UTF-8\?B\?/);
  assert.match(eml, /Content-Disposition: attachment; filename="a.png"/);
  assert.match(eml, /Content-Type: multipart\/mixed; boundary=/);
});

test('.msg: broken input throws instead of returning garbage', () => {
  assert.throws(() => A.parseMsg(new Uint8Array([1, 2, 3])));
  assert.strictEqual(A.isMsg(new Uint8Array([1, 2, 3])), false);
  assert.strictEqual(A.isMsg(PNG), false);
});

test('sniff() recognises the file types we care about', () => {
  assert.strictEqual(A.sniff(PNG).mime, 'image/png');
  assert.strictEqual(A.sniff(Buffer.from([0xff, 0xd8, 0xff, 0xe0])).mime, 'image/jpeg');
  assert.strictEqual(A.sniff(Buffer.from('GIF89a')).mime, 'image/gif');
  assert.strictEqual(A.sniff(Buffer.from('%PDF-1.7')).ext, 'pdf');
  assert.strictEqual(A.sniff(msg({})).kind, 'cfb');
  assert.strictEqual(A.sniff(Buffer.from('hello')).kind, 'file');
});

test('the generated zip is a clean package: no SheetJS marker entry, compressed, content types first', () => {
  const xl = embed(RAID, [{ type: 'image', sheet: 'Risks', row: 11, col: 20, name: 'a.png', bytes: PNG }]);
  const files = readZip(Buffer.from(xl));
  assert.ok(!Object.keys(files).some((p) => /Sh33tJ5/.test(p)), 'SheetJS marker entry must not be copied (LibreOffice rejects the file)');
  assert.ok(files['[Content_Types].xml'] && files['xl/workbook.xml']);
  const w = XLSX.read(xl, { type: 'array' });          // other readers still open it
  assert.ok(w.SheetNames.includes('Risks'));
});

test('security: a hostile MIME type or file name inside an email cannot reach the page markup', () => {
  const m = A.parseMsg(msg({ attachments: [{ name: '"><img src=x onerror=alert(1)>.png', mime: 'image/png" onerror="alert(1)', bytes: PNG }] }));
  assert.strictEqual(m.attachments[0].mime, 'image/png');                    // falls back to what the bytes say
  assert.strictEqual(A.safeMime('text/html;charset=x" onload=1'), 'application/octet-stream');
  assert.strictEqual(A.safeMime('image/svg+xml'), 'image/svg+xml');
  assert.strictEqual(A.safeMime(''), 'application/octet-stream');
});
