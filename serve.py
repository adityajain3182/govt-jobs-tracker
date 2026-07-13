#!/usr/bin/env python3
"""Local dev server for the govt-jobs-tracker static site.

Run:  python serve.py            (or: npm run serve)
Then open http://localhost:8000

Serves the repo it lives in, maps the clean URL /applications -> applications.html,
and disables caching so edits show on a plain refresh.
"""
import http.server
import os
import socketserver

ROOT = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get("PORT", "8000"))


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **k):
        super().__init__(*a, directory=ROOT, **k)

    def translate_path(self, path):
        clean = path.split("?", 1)[0].split("#", 1)[0].rstrip("/")
        if clean == "/applications":
            return os.path.join(ROOT, "applications.html")
        return super().translate_path(path)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()


with socketserver.TCPServer(("127.0.0.1", PORT), Handler) as httpd:
    print(f"Serving {ROOT}\n  http://localhost:{PORT}  (Ctrl+C to stop)")
    httpd.serve_forever()
