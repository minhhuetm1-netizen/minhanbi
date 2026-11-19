import os
import sys
import mimetypes
import urllib.parse

# Ensure root directory is on sys.path
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from server import MinhAnRequestHandler

class handler(MinhAnRequestHandler):
    """
    Vercel Serverless Function entry point:
    Handles API endpoints and serves static assets (index.html, app.js, chart.umd.js)
    with zero configuration conflicts on Vercel.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=root_dir, **kwargs)

    def _get_clean_path(self):
        p = self.path
        if '?' in p:
            p = p.split('?', 1)[0]
        # If Vercel rewrites pass the original matched path in headers
        if p in ('/api/index.py', '/api/index', '/api'):
            orig = self.headers.get('x-matched-path') or self.headers.get('x-forwarded-uri') or '/'
            if '?' in orig:
                orig = orig.split('?', 1)[0]
            p = orig
        return p

    def do_GET(self):
        req_path = self._get_clean_path()

        # If requesting root homepage or index.html
        if req_path in ('/', '/index.html', ''):
            html_file = os.path.join(root_dir, 'index.html')
            if os.path.exists(html_file):
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                with open(html_file, 'rb') as f:
                    content = f.read()
                self.send_header('Content-Length', str(len(content)))
                self.end_headers()
                self.wfile.write(content)
                return

        # If requesting a static file at root or data/ (e.g. app.js, chart.umd.js, favicon.ico)
        filename = req_path.lstrip('/')
        file_path = os.path.join(root_dir, filename)
        if os.path.isfile(file_path) and not filename.endswith('.py') and not filename.startswith('.'):
            mime_type, _ = mimetypes.guess_type(file_path)
            if not mime_type:
                mime_type = 'application/octet-stream'
            if mime_type.startswith('text/') or mime_type in ('application/javascript', 'application/json'):
                mime_type += '; charset=utf-8'
            self.send_response(200)
            self.send_header('Content-Type', mime_type)
            with open(file_path, 'rb') as f:
                content = f.read()
            self.send_header('Content-Length', str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return

        # Otherwise delegate to API routing in MinhAnRequestHandler
        self.path = req_path
        super().do_GET()

    def do_POST(self):
        self.path = self._get_clean_path()
        super().do_POST()
