from __future__ import annotations

import json
from io import BytesIO
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs, urlparse

from sermon_engine import render_infographic, render_slides_pdf, slugify, structure_from_payload


def generate_export(payload: dict, kind: str) -> tuple[bytes, str, str]:
    structure = structure_from_payload(payload)
    slug = slugify(structure.titulo or "sermao")
    buffer = BytesIO()

    if kind == "infografico":
        render_infographic(structure).save(buffer, "PNG", optimize=True)
        return buffer.getvalue(), "image/png", f"{slug}_infografico.png"
    if kind == "pdf":
        render_slides_pdf(structure, buffer)
        return buffer.getvalue(), "application/pdf", f"{slug}_slides_cards.pdf"
    raise ValueError("Tipo de arquivo inválido")


class handler(BaseHTTPRequestHandler):
    def do_POST(self) -> None:
        length = int(self.headers.get("Content-Length", "0"))
        payload = json.loads(self.rfile.read(length).decode("utf-8"))
        kind = parse_qs(urlparse(self.path).query).get("kind", [""])[0]
        try:
            data, content_type, filename = generate_export(payload, kind)
        except ValueError:
            self.send_error(HTTPStatus.BAD_REQUEST, "Tipo de arquivo inválido")
            return
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)
