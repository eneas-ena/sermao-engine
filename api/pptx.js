import pptxgen from "pptxgenjs";

const NAVY = "041A2E";
const NAVY_2 = "082B48";
const GOLD = "DEA94A";
const CREAM = "F7F2E8";
const WHITE = "FFFFFF";
const INK = "0F1A2B";
const MUTED = "546273";

function addShape(slide, pptx, type, x, y, w, h, fill, line = fill, radius = 0) {
  const shape = radius ? pptx.ShapeType.roundRect : type;
  slide.addShape(shape, {
    x, y, w, h,
    fill: { color: fill },
    line: { color: line, width: line === fill ? 0 : 1.5 },
    radius,
  });
}

function addText(slide, value, x, y, w, h, options = {}) {
  slide.addText(String(value ?? ""), {
    x, y, w, h,
    fontFace: options.fontFace ?? "Aptos",
    fontSize: options.fontSize ?? 22,
    bold: options.bold ?? false,
    color: options.color ?? INK,
    align: options.align ?? "left",
    valign: options.valign ?? "top",
    margin: options.margin ?? 0,
    breakLine: false,
    fit: "shrink",
  });
}

function addHeader(slide, pptx, data, heading, slideTitleSize) {
  addShape(slide, pptx, pptx.ShapeType.rect, 0, 0, 13.333, 1.08, NAVY);
  addShape(slide, pptx, pptx.ShapeType.rect, 0, 1.05, 13.333, 0.03, GOLD);
  addText(slide, heading.toUpperCase(), 0.55, 0.22, 10.7, 0.48, {
    fontSize: slideTitleSize, bold: true, color: GOLD,
  });
  addText(slide, (data.titulo ?? "").toUpperCase(), 0.56, 0.76, 11.2, 0.2, {
    fontSize: 12, bold: true, color: WHITE,
  });
}

function addFooter(slide) {
  addText(slide, "Infográfico de sermão | conteúdo e elementos editáveis", 0.56, 7.05, 8.3, 0.2, {
    fontSize: 9, color: MUTED,
  });
}

function addTitleSlide(pptx, data, style) {
  const slide = pptx.addSlide();
  slide.background = { color: NAVY };
  addShape(slide, pptx, pptx.ShapeType.rect, 0, 0, 0.16, 7.5, GOLD);
  const hasImage = typeof data.imagem_capa === "string" && data.imagem_capa.startsWith("data:image/");
  const textWidth = hasImage ? 7.1 : 10.9;
  addText(slide, (data.titulo ?? "").toUpperCase(), 0.7, 0.82, textWidth, 2.25, {
    fontSize: style.titulo, bold: true, color: WHITE,
  });
  addText(slide, (data.subtitulo ?? "").toUpperCase(), 0.72, 3.25, textWidth, 1.05, {
    fontSize: Math.max(22, Math.round(style.titulo * 0.48)), bold: true, color: GOLD,
  });
  if (hasImage) {
    slide.addImage({ data: data.imagem_capa, x: 8.25, y: 0.62, w: 4.5, h: 5.7 });
  }
  const bibleWidth = hasImage ? 7.1 : 11.25;
  addShape(slide, pptx, pptx.ShapeType.roundRect, 0.7, 5.12, bibleWidth, 1.45, NAVY_2, GOLD, 0.08);
  addText(slide, `TEXTO: ${data.texto_biblico?.referencia ?? ""}`.toUpperCase(), 1.0, 5.36, bibleWidth - 0.6, 0.35, {
    fontSize: 21, bold: true, color: GOLD,
  });
  addText(slide, data.texto_biblico?.texto ?? "", 1.0, 5.8, bibleWidth - 0.6, 0.55, {
    fontSize: 17, color: WHITE,
  });
}

function addTextSlide(pptx, data, style, heading, content) {
  const slide = pptx.addSlide();
  slide.background = { color: CREAM };
  addHeader(slide, pptx, data, heading, style.titulo_slide);
  addShape(slide, pptx, pptx.ShapeType.roundRect, 0.75, 1.6, 11.85, 4.72, WHITE, GOLD, 0.08);
  addShape(slide, pptx, pptx.ShapeType.rect, 0.75, 1.6, 0.12, 4.72, GOLD);
  addText(slide, content, 1.28, 2.08, 10.7, 3.6, {
    fontSize: Math.min(style.corpo + 6, 32), color: INK,
  });
  addFooter(slide);
}

function addCardSlide(pptx, data, style, card) {
  const slide = pptx.addSlide();
  slide.background = { color: NAVY };
  addShape(slide, pptx, pptx.ShapeType.roundRect, 0.82, 0.72, 11.68, 5.92, CREAM, GOLD, 0.12);
  addShape(slide, pptx, pptx.ShapeType.ellipse, 1.28, 1.2, 1.04, 1.04, GOLD);
  addText(slide, card.numero ?? "", 1.28, 1.35, 1.04, 0.55, {
    fontSize: 42, bold: true, color: NAVY, align: "center", valign: "mid",
  });
  addText(slide, (card.titulo ?? "").toUpperCase(), 2.7, 1.18, 9.1, 1.0, {
    fontSize: style.titulo_slide, bold: true, color: NAVY,
  });
  if (card.referencia) {
    addShape(slide, pptx, pptx.ShapeType.roundRect, 2.7, 2.32, 2.5, 0.43, NAVY, NAVY, 0.06);
    addText(slide, card.referencia, 2.9, 2.4, 2.1, 0.22, {
      fontSize: 15, bold: true, color: WHITE,
    });
  }
  const summaryLength = String(card.resumo ?? "").length;
  const summarySize = summaryLength > 220 ? 18 : summaryLength > 150 ? 20 : Math.min(style.corpo + 2, 24);
  addText(slide, card.resumo ?? "", 1.42, 3.08, 10.5, 1.92, {
    fontSize: summarySize, color: INK,
  });
  if (card.frase) {
    addShape(slide, pptx, pptx.ShapeType.roundRect, 1.42, 5.28, 10.5, 0.9, NAVY, NAVY, 0.08);
    addText(slide, `“${card.frase}”`, 1.8, 5.5, 9.75, 0.45, {
      fontSize: Math.min(style.corpo + 3, 27), bold: true, color: GOLD, align: "center", valign: "mid",
    });
  }
}

export async function buildPptx(data) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Infográficos de Sermões";
  pptx.subject = data.titulo ?? "Sermão";
  pptx.title = data.titulo ?? "Sermão";
  pptx.company = "Uso pessoal";
  pptx.lang = "pt-BR";
  pptx.theme = {
    headFontFace: "Aptos Display",
    bodyFontFace: "Aptos",
    lang: "pt-BR",
  };

  const style = {
    titulo: Math.max(36, Math.min(72, Number(data.estilo?.titulo ?? 54))),
    titulo_slide: Math.max(26, Math.min(48, Number(data.estilo?.titulo_slide ?? 34))),
    corpo: Math.max(16, Math.min(32, Number(data.estilo?.corpo ?? 22))),
  };

  addTitleSlide(pptx, data, style);
  addTextSlide(pptx, data, style, "Introdução", data.introducao ?? "");
  for (const card of data.cards ?? []) addCardSlide(pptx, data, style, card);
  addTextSlide(pptx, data, style, "Aplicação prática", data.aplicacao ?? "");
  addTextSlide(pptx, data, style, "Desafio", data.desafio ?? "");
  addTextSlide(pptx, data, style, "Apelo", data.apelo ?? "");
  return pptx.write({ outputType: "nodebuffer" });
}

async function readJson(req) {
  if (req.body && typeof req.body === "object" && !Buffer.isBuffer(req.body)) return req.body;
  if (typeof req.body === "string") return JSON.parse(req.body);
  const chunks = [];
  for await (const chunk of req) chunks.push(Buffer.from(chunk));
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.statusCode = 405;
    res.end("Método não permitido");
    return;
  }
  try {
    const data = await readJson(req);
    const buffer = await buildPptx(data);
    res.statusCode = 200;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", "attachment; filename=slides_editaveis.pptx");
    res.setHeader("Cache-Control", "no-store");
    res.end(buffer);
  } catch (error) {
    console.error(error);
    res.statusCode = 500;
    res.end("Não foi possível gerar o PowerPoint.");
  }
}
