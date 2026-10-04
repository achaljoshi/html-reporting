#!/usr/bin/env node
// Tiny dependency-free static server for QA / development (the shipped app needs no server — index.html works from file://).
// usage: node tools/serve.js          (PORT env overrides the default 8765)
const http = require('http'), fs = require('fs'), path = require('path');
const root = path.resolve(__dirname, '..');
const port = parseInt(process.env.PORT, 10) || 8765;
const TYPES = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.svg': 'image/svg+xml', '.png': 'image/png', '.jpg': 'image/jpeg', '.ico': 'image/x-icon', '.md': 'text/plain; charset=utf-8', '.txt': 'text/plain; charset=utf-8',
};
const server = http.createServer((req, res) => {
  const headers = { 'Cache-Control': 'no-store' };
  let rel;
  try { rel = decodeURIComponent(new URL(req.url, 'http://localhost').pathname); } catch (e) { res.writeHead(400, headers); return res.end('Bad request'); }
  let file = path.join(root, rel);
  if (file !== root && !file.startsWith(root + path.sep)) { res.writeHead(403, headers); return res.end('Forbidden'); }
  try { if (fs.statSync(file).isDirectory()) file = path.join(file, 'index.html'); } catch (e) { /* handled below */ }
  fs.readFile(file, (err, buf) => {
    if (err) { res.writeHead(404, Object.assign({ 'Content-Type': 'text/plain' }, headers)); return res.end('Not found'); }
    res.writeHead(200, Object.assign({ 'Content-Type': TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream' }, headers));
    res.end(req.method === 'HEAD' ? undefined : buf);
  });
});
server.listen(port, () => console.log('ATS dashboard serving ' + root + ' at http://localhost:' + port));
