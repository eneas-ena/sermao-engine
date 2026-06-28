from __future__ import annotations

import cgi
import json
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler

from sermon_engine import analyze_sermon, extract_docx_text


def analyze_submission(text: str, docx_bytes: bytes | None = None) -> dict:
    if docx_bytes:
        text = extract_docx_text(docx_bytes)
    return {"text": text, "structure": analyze_sermon(text).to_dict()}


class handler(BaseHTTPRequestHandler):
    def do_POST(self) -> None:
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
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
