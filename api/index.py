from __future__ import annotations

import cgi
import json
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from api.analyze import analyze_submission
from api.export import generate_export


STATIC_ROOT = Path(__file__).resolve().parents[1] / "static"
STATIC_FILES = {
    "": ("index.html", "text/html; charset=utf-8"),
    "/": ("index.html", "text/html; charset=utf-8"),
    "/static/index.html": ("index.html", "text/html; charset=utf-8"),
    "/static/styles.css": ("styles.css", "text/css; charset=utf-8"),
    "/static/app.js": ("app.js", "application/javascript; charset=utf-8"),
}


class handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        path = urlparse(self.path).path
        static_file = STATIC_FILES.get(path)
        if static_file is None:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        filename, content_type = static_file
        self._send((STATIC_ROOT / filename).read_bytes(), content_type)

    def do_POST(self) -> None:
        path = urlparse(self.path).path.rstrip("/")
        if path == "/api/analyze":
            self._analyze()
            return
        if path == "/api/export":
            self._export()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def _analyze(self) -> None:
        form = cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
            },
        )
        text = form.getfirst("text", "").strip()
        file_item = form["docx"] if "docx" in form else None
        docx_bytes = file_item.file.read() if file_item is not None and getattr(file_item, "filename", "") else None
        data = json.dumps(analyze_submission(text, docx_bytes), ensure_ascii=False).encode("utf-8")
        self._send(data, "application/json; charset=utf-8")

    def _export(self) -> None:
        length = int(self.headers.get("Content-Length", "0"))
        payload = json.loads(self.rfile.read(length).decode("utf-8"))
        kind = parse_qs(urlparse(self.path).query).get("kind", [""])[0]
        try:
            data, content_type, filename = generate_export(payload, kind)
        except ValueError:
            self.send_error(HTTPStatus.BAD_REQUEST, "Tipo de arquivo inválido")
            return
        self._send(data, content_type, filename)

    def _send(self, data: bytes, content_type: str, filename: str | None = None) -> None:
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        if filename:
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(data)
