#!/usr/bin/env python3
"""Local QA-only static server with caching disabled (the shipped app needs no server)."""
import sys, http.server, socketserver, os
port = int(sys.argv[1]) if len(sys.argv) > 1 else 8743
root = sys.argv[2] if len(sys.argv) > 2 else os.getcwd()
os.chdir(root)
class H(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        super().end_headers()
    def do_POST(self):
        # QA-only: lets the test browser save screenshots to $ATS_SHOT_DIR
        d = os.environ.get("ATS_SHOT_DIR")
        if not d or not self.path.startswith("/__save/"):
            self.send_response(404); self.end_headers(); return
        name = os.path.basename(self.path[len("/__save/"):])
        n = int(self.headers.get("Content-Length", 0))
        with open(os.path.join(d, name), "wb") as f: f.write(self.rfile.read(n))
        self.send_response(200); self.end_headers(); self.wfile.write(b"ok")
    def log_message(self, *a): pass
socketserver.TCPServer.allow_reuse_address = True
with socketserver.TCPServer(("", port), H) as httpd:
    httpd.serve_forever()
