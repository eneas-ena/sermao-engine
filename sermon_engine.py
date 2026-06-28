from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from io import BytesIO
from pathlib import Path
from typing import Any

from docx import Document
from PIL import Image, ImageDraw, ImageFilter, ImageFont
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas


SLIDE_W = 1920
SLIDE_H = 1080
NAVY = (4, 26, 46)
NAVY_2 = (8, 43, 72)
GOLD = (222, 169, 74)
CREAM = (247, 242, 232)
WHITE = (255, 255, 255)
INK = (15, 26, 43)
FONT_DIR = Path(__file__).resolve().parent / "assets"


@dataclass
class BibleText:
    referencia: str = ""
    texto: str = ""


@dataclass
class SermonCard:
    numero: int
    titulo: str
    referencia: str
    resumo: str
    frase: str


@dataclass
class PresentationStyle:
    titulo: int = 54
    titulo_slide: int = 34
    corpo: int = 22


@dataclass
class SermonStructure:
    titulo: str
    subtitulo: str
    texto_biblico: BibleText
    introducao: str
    cards: list[SermonCard]
    aplicacao: str
    desafio: str
    apelo: str
    estilo: PresentationStyle
    imagem_capa: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def extract_docx_text(file_bytes: bytes) -> str:
    doc = Document(BytesIO(file_bytes))
    lines: list[str] = []
    for paragraph in doc.paragraphs:
        text = " ".join(paragraph.text.split())
        if text:
            lines.append(text)
    return "\n".join(lines)


def analyze_sermon(text: str) -> SermonStructure:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return empty_structure()

    title = lines[0]
    bible_ref = ""
    bible_text = ""
    subtitle = ""

    for i, line in enumerate(lines[1:8], start=1):
        if line.upper().startswith("TEXTO:"):
            bible_ref = line.split(":", 1)[1].strip()
            if i + 1 < len(lines):
                bible_text = lines[i + 1].strip()
            break

    intro = extract_between(lines, "INTRODUÇÃO", r"^\d+\.\s+")
    if intro:
        subtitle = first_sentence(intro)
    else:
        subtitle = "Mensagem bíblica organizada para ensino, decisão e aplicação."

    sections = split_numbered_sections(lines)
    cards = [make_card(i + 1, section_title, section_lines) for i, (section_title, section_lines) in enumerate(sections[:8])]

    conclusion = extract_after_heading(lines, "CONCLUSÃO", stop_headings=["DESAFIO FINAL:", "APELO:"])
    challenge = extract_after_heading(lines, "DESAFIO FINAL:", stop_headings=["APELO:"])
    appeal = extract_after_heading(lines, "APELO:", stop_headings=[])

    if not challenge:
        challenge = "Examine sua fé, responda em obediência e dê o próximo passo diante de Deus."
    if not appeal:
        appeal = conclusion or "Responda hoje com fé, obediência e confiança na graça de Deus."

    return SermonStructure(
        titulo=title,
        subtitulo=subtitle,
        texto_biblico=BibleText(referencia=bible_ref, texto=bible_text),
        introducao=shorten(intro, 420),
        cards=cards,
        aplicacao=build_application(cards, conclusion),
        desafio=shorten(challenge, 420),
        apelo=shorten(appeal, 420),
        estilo=PresentationStyle(),
    )


def empty_structure() -> SermonStructure:
    return SermonStructure(
        titulo="Novo infográfico bíblico",
        subtitulo="",
        texto_biblico=BibleText(),
        introducao="",
        cards=[],
        aplicacao="",
        desafio="",
        apelo="",
        estilo=PresentationStyle(),
    )


def extract_between(lines: list[str], start_heading: str, stop_pattern: str) -> str:
    collecting = False
    collected: list[str] = []
    stop_re = re.compile(stop_pattern)
    for line in lines:
        if line.upper() == start_heading.upper():
            collecting = True
            continue
        if collecting and stop_re.match(line):
            break
        if collecting:
            collected.append(line)
    return " ".join(collected)


def extract_after_heading(lines: list[str], heading: str, stop_headings: list[str]) -> str:
    collecting = False
    collected: list[str] = []
    heading_upper = heading.upper()
    stop_set = {item.upper() for item in stop_headings}
    for line in lines:
        if line.upper() == heading_upper:
            collecting = True
            continue
        if collecting and line.upper() in stop_set:
            break
        if collecting:
            collected.append(line)
    return " ".join(collected)


def split_numbered_sections(lines: list[str]) -> list[tuple[str, list[str]]]:
    sections: list[tuple[str, list[str]]] = []
    current_title = ""
    current_lines: list[str] = []
    marker = re.compile(r"^(\d+)\.\s+(.+)")

    for line in lines:
        match = marker.match(line)
        if match:
            if current_title:
                sections.append((current_title, current_lines))
            current_title = match.group(2).strip()
            current_lines = []
            continue
        if current_title:
            if line.upper() in {"CONCLUSÃO", "DESAFIO FINAL:", "APELO:"}:
                break
            current_lines.append(line)

    if current_title:
        sections.append((current_title, current_lines))
    return sections


def make_card(number: int, title: str, lines: list[str]) -> SermonCard:
    joined = " ".join(lines)
    phrase = ""
    phrase_match = re.search(r"Frase de Impacto:\s*(.+?)(?:\s+Transição:|$)", joined, re.IGNORECASE)
    if phrase_match:
        phrase = phrase_match.group(1).strip().strip('"“”')

    reference = ""
    ref_match = re.search(r"\(([1-3]?\s?[A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-Za-zÁÉÍÓÚáéíóúâêôãõç]+\s+\d+:\d+(?:-\d+)?)\)", joined)
    if ref_match:
        reference = ref_match.group(1)

    cleaned = re.sub(r"Frase de Impacto:.+?(?=Transição:|$)", "", joined, flags=re.IGNORECASE)
    cleaned = re.sub(r"Transição:.+$", "", cleaned, flags=re.IGNORECASE).strip()
    summary = build_summary(cleaned)

    return SermonCard(
        numero=number,
        titulo=title.title(),
        referencia=reference,
        resumo=shorten(summary, 280),
        frase=shorten(phrase, 150),
    )


def build_summary(text: str) -> str:
    sentences = split_sentences(text)
    meaningful = [s for s in sentences if len(s) > 45]
    return " ".join(meaningful[:2] or sentences[:2])


def build_application(cards: list[SermonCard], conclusion: str) -> str:
    if conclusion:
        return shorten(conclusion, 360)
    if cards:
        return "Responda à Palavra com fé, obediência e compromisso público diante de Cristo."
    return ""


def split_sentences(text: str) -> list[str]:
    parts = re.split(r"(?<=[.!?])\s+", text)
    return [part.strip() for part in parts if part.strip()]


def first_sentence(text: str) -> str:
    sentences = split_sentences(text)
    return sentences[0] if sentences else ""


def shorten(text: str, limit: int) -> str:
    text = " ".join(text.split())
    if len(text) <= limit:
        return text
    return text[: limit - 1].rsplit(" ", 1)[0] + "…"


def structure_from_payload(payload: dict[str, Any]) -> SermonStructure:
    style_payload = payload.get("estilo", {})
    return SermonStructure(
        titulo=payload.get("titulo", ""),
        subtitulo=payload.get("subtitulo", ""),
        texto_biblico=BibleText(**payload.get("texto_biblico", {})),
        introducao=payload.get("introducao", ""),
        cards=[SermonCard(**card) for card in payload.get("cards", [])],
        aplicacao=payload.get("aplicacao", ""),
        desafio=payload.get("desafio", ""),
        apelo=payload.get("apelo", ""),
        estilo=PresentationStyle(
            titulo=clamp_int(style_payload.get("titulo", 54), 36, 72),
            titulo_slide=clamp_int(style_payload.get("titulo_slide", 34), 26, 48),
            corpo=clamp_int(style_payload.get("corpo", 22), 16, 32),
        ),
        imagem_capa=payload.get("imagem_capa", ""),
    )


def clamp_int(value: Any, minimum: int, maximum: int) -> int:
    try:
        number = int(value)
    except (TypeError, ValueError):
        number = minimum
    return max(minimum, min(maximum, number))


def render_infographic(structure: SermonStructure) -> Image.Image:
    image = Image.new("RGB", (SLIDE_W, SLIDE_H), NAVY)
    draw = ImageDraw.Draw(image)
    add_background(image)

    title_font = font(56, bold=True)
    subtitle_font = font(28, bold=True)
    body_font = font(24)
    small_font = font(20)

    draw_wrapped(draw, structure.titulo.upper(), (70, 42), 930, title_font, WHITE, line_gap=8, max_lines=2)
    draw_wrapped(draw, structure.subtitulo.upper(), (70, 165), 920, subtitle_font, GOLD, line_gap=6, max_lines=1)

    bible_box = (1060, 42, 1810, 190)
    rounded(draw, bible_box, NAVY_2, GOLD, 18, 2)
    draw.text((1080, 70), f"TEXTO: {structure.texto_biblico.referencia}".upper(), font=font(30, bold=True), fill=GOLD)
    draw_wrapped(draw, structure.texto_biblico.texto, (1080, 112), 690, body_font, WHITE, line_gap=8, max_lines=3)

    intro_box = (58, 220, 1862, 315)
    rounded(draw, intro_box, (5, 35, 60), GOLD, 14, 2)
    draw.text((92, 236), "INTRODUÇÃO", font=font(28, bold=True), fill=GOLD)
    draw_wrapped(draw, structure.introducao, (300, 235), 1500, small_font, WHITE, line_gap=7, max_lines=3)

    render_card_grid(draw, structure.cards[:4])
    render_footer(draw, structure)

    return image


def add_background(image: Image.Image) -> None:
    overlay = Image.new("RGBA", image.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    for i in range(0, SLIDE_W, 24):
        alpha = int(16 * (i / SLIDE_W))
        draw.line((i, 0, i - 500, SLIDE_H), fill=(255, 255, 255, alpha), width=1)
    draw.ellipse((700, 210, 1220, 730), fill=(222, 169, 74, 32))
    draw.ellipse((835, 335, 1085, 585), outline=(222, 169, 74, 180), width=8)
    draw.line((960, 355, 960, 565), fill=(255, 255, 255, 210), width=20)
    draw.line((885, 430, 1035, 430), fill=(255, 255, 255, 210), width=18)
    blurred = overlay.filter(ImageFilter.GaussianBlur(1))
    image.paste(Image.alpha_composite(image.convert("RGBA"), blurred).convert("RGB"))


def render_card_grid(draw: ImageDraw.ImageDraw, cards: list[SermonCard]) -> None:
    positions = [
        (58, 345, 468, 765),
        (510, 345, 920, 765),
        (962, 345, 1372, 765),
        (1414, 345, 1810, 765),
    ]

    for card, box in zip(cards, positions):
        rounded(draw, box, CREAM, GOLD, 16, 3)
        x1, y1, x2, _ = box
        badge = (x1 + 20, y1 + 20, x1 + 92, y1 + 92)
        draw.ellipse(badge, fill=GOLD)
        draw.text((x1 + 43, y1 + 27), str(card.numero), font=font(44, bold=True), fill=NAVY)
        draw_wrapped(draw, card.titulo.upper(), (x1 + 105, y1 + 22), x2 - x1 - 125, font(23, bold=True), NAVY, line_gap=5, max_lines=2)
        if card.referencia:
            ref_y = y1 + 102
            draw.rounded_rectangle((x1 + 105, ref_y, min(x1 + 310, x2 - 25), ref_y + 38), radius=10, fill=NAVY)
            draw.text((x1 + 122, ref_y + 7), card.referencia, font=font(18, bold=True), fill=WHITE)
        draw_wrapped(draw, card.resumo, (x1 + 32, y1 + 155), x2 - x1 - 64, font(20), INK, line_gap=7, max_lines=6)
        if card.frase:
            phrase_box = (x1 + 22, y2(box) - 82, x2 - 22, y2(box) - 20)
            draw.rounded_rectangle(phrase_box, radius=12, fill=NAVY)
            draw_wrapped(draw, f"“{card.frase}”", (phrase_box[0] + 18, phrase_box[1] + 10), phrase_box[2] - phrase_box[0] - 36, font(18, bold=True), GOLD, line_gap=4, max_lines=2)


def y2(box: tuple[int, int, int, int]) -> int:
    return box[3]


def render_footer(draw: ImageDraw.ImageDraw, structure: SermonStructure) -> None:
    footer_top = 795
    boxes = [
        ("APLICAÇÃO", structure.aplicacao, (58, footer_top, 610, 1015)),
        ("DESAFIO", structure.desafio, (684, footer_top, 1236, 1015)),
        ("APELO", structure.apelo, (1262, footer_top, 1810, 1015)),
    ]
    for title, text, box in boxes:
        rounded(draw, box, (5, 35, 60), GOLD, 16, 2)
        draw.text((box[0] + 24, box[1] + 20), title, font=font(26, bold=True), fill=GOLD)
        draw_wrapped(draw, text, (box[0] + 24, box[1] + 62), box[2] - box[0] - 48, font(20), WHITE, line_gap=7, max_lines=5)


def render_slides_pdf(structure: SermonStructure, pdf_path: Path | BytesIO) -> None:
    pages = [render_title_slide(structure)]
    pages.append(render_text_slide("INTRODUÇÃO", structure.introducao, structure))
    for card in structure.cards:
        pages.append(render_card_slide(structure, card))
    pages.append(render_text_slide("APLICAÇÃO PRÁTICA", structure.aplicacao, structure))
    pages.append(render_text_slide("DESAFIO", structure.desafio, structure))
    pages.append(render_text_slide("APELO", structure.apelo, structure))

    destination = str(pdf_path) if isinstance(pdf_path, Path) else pdf_path
    pdf = canvas.Canvas(destination, pagesize=(SLIDE_W, SLIDE_H))
    for page in pages:
        buffer = BytesIO()
        page.save(buffer, "PNG")
        buffer.seek(0)
        pdf.drawImage(ImageReader(buffer), 0, 0, width=SLIDE_W, height=SLIDE_H)
        pdf.showPage()
    pdf.save()


def render_title_slide(structure: SermonStructure) -> Image.Image:
    image = Image.new("RGB", (SLIDE_W, SLIDE_H), NAVY)
    draw = ImageDraw.Draw(image)
    add_background(image)
    draw_wrapped(draw, structure.titulo.upper(), (90, 120), 1300, font(86, bold=True), WHITE, line_gap=12, max_lines=3)
    draw_wrapped(draw, structure.subtitulo.upper(), (95, 385), 1200, font(34, bold=True), GOLD, line_gap=8, max_lines=2)
    rounded(draw, (1010, 685, 1810, 900), NAVY_2, GOLD, 18, 3)
    draw.text((1050, 725), f"TEXTO: {structure.texto_biblico.referencia}".upper(), font=font(34, bold=True), fill=GOLD)
    draw_wrapped(draw, structure.texto_biblico.texto, (1050, 782), 700, font(28), WHITE, line_gap=10, max_lines=3)
    return image


def render_text_slide(title: str, text: str, structure: SermonStructure) -> Image.Image:
    image = Image.new("RGB", (SLIDE_W, SLIDE_H), CREAM)
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, SLIDE_W, 150), fill=NAVY)
    draw.text((80, 48), title.upper(), font=font(48, bold=True), fill=GOLD)
    draw.text((80, 108), structure.titulo.upper(), font=font(20, bold=True), fill=WHITE)
    rounded(draw, (110, 230, 1810, 870), WHITE, GOLD, 20, 3)
    draw_wrapped(draw, text, (180, 300), 1560, font(34), INK, line_gap=16, max_lines=11)
    return image


def render_card_slide(structure: SermonStructure, card: SermonCard) -> Image.Image:
    image = Image.new("RGB", (SLIDE_W, SLIDE_H), NAVY)
    draw = ImageDraw.Draw(image)
    add_background(image)
    rounded(draw, (130, 120, 1790, 930), CREAM, GOLD, 26, 4)
    draw.ellipse((190, 190, 330, 330), fill=GOLD)
    draw.text((240, 205), str(card.numero), font=font(76, bold=True), fill=NAVY)
    draw_wrapped(draw, card.titulo.upper(), (390, 190), 1280, font(56, bold=True), NAVY, line_gap=10, max_lines=2)
    if card.referencia:
        draw.rounded_rectangle((390, 350, 710, 405), radius=14, fill=NAVY)
        draw.text((420, 361), card.referencia, font=font(28, bold=True), fill=WHITE)
    draw_wrapped(draw, card.resumo, (210, 475), 1500, font(38), INK, line_gap=18, max_lines=5)
    if card.frase:
        draw.rounded_rectangle((210, 785, 1710, 890), radius=20, fill=NAVY)
        draw_wrapped(draw, f"“{card.frase}”", (250, 812), 1420, font(34, bold=True), GOLD, line_gap=10, max_lines=2)
    return image


def rounded(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], fill: tuple[int, int, int], outline: tuple[int, int, int], radius: int, width: int) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


def draw_wrapped(
    draw: ImageDraw.ImageDraw,
    text: str,
    xy: tuple[int, int],
    max_width: int,
    chosen_font: ImageFont.ImageFont,
    fill: tuple[int, int, int],
    line_gap: int = 6,
    max_lines: int | None = None,
) -> None:
    if not text:
        return
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        trial = f"{current} {word}".strip()
        if text_width(draw, trial, chosen_font) <= max_width or not current:
            current = trial
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1].rstrip(".,;:") + "…"
    x, y = xy
    line_height = chosen_font.size + line_gap if hasattr(chosen_font, "size") else 28
    for line in lines:
        draw.text((x, y), line, font=chosen_font, fill=fill)
        y += line_height


def text_width(draw: ImageDraw.ImageDraw, text: str, chosen_font: ImageFont.ImageFont) -> int:
    bbox = draw.textbbox((0, 0), text, font=chosen_font)
    return bbox[2] - bbox[0]


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        str(FONT_DIR / ("DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf")),
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial Bold.ttf" if bold else "/Library/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def slugify(value: str) -> str:
    value = value.lower()
    replacements = {
        "á": "a", "à": "a", "â": "a", "ã": "a",
        "é": "e", "ê": "e",
        "í": "i",
        "ó": "o", "ô": "o", "õ": "o",
        "ú": "u",
        "ç": "c",
    }
    for old, new in replacements.items():
        value = value.replace(old, new)
    value = re.sub(r"[^a-z0-9]+", "-", value).strip("-")
    return value[:80] or "sermao"

