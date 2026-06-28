let structure = null;
let generatedUrls = [];

const $ = (id) => document.getElementById(id);

const fields = {
  title: $("titleInput"),
  subtitle: $("subtitleInput"),
  ref: $("refInput"),
  verse: $("verseInput"),
  intro: $("introInput"),
  application: $("applicationInput"),
  challenge: $("challengeInput"),
  appeal: $("appealInput"),
  titleSize: $("titleSizeInput"),
  slideTitleSize: $("slideTitleSizeInput"),
  bodySize: $("bodySizeInput"),
};

$("analyzeBtn").addEventListener("click", analyze);
$("exportBtn").addEventListener("click", exportFiles);
$("addCardBtn").addEventListener("click", addCard);
$("clearBtn").addEventListener("click", clearWorkspace);
$("projectInput").addEventListener("change", loadProject);

async function analyze() {
  setStatus("Analisando...");
  const form = new FormData();
  const file = $("docxInput").files[0];
  if (file) form.append("docx", file);
  form.append("text", $("sermonText").value);

  const response = await fetch("/api/analyze", { method: "POST", body: form });
  if (!response.ok) {
    setStatus("Não foi possível analisar o sermão");
    return;
  }
  const data = await response.json();
  $("sermonText").value = data.text || $("sermonText").value;
  structure = data.structure;
  populateEditor();
  $("exportBtn").disabled = false;
  $("addCardBtn").disabled = false;
  $("clearBtn").disabled = false;
  setStatus("Pronto para revisão");
}

function populateEditor() {
  fields.title.value = structure.titulo || "";
  fields.subtitle.value = structure.subtitulo || "";
  fields.ref.value = structure.texto_biblico?.referencia || "";
  fields.verse.value = structure.texto_biblico?.texto || "";
  fields.intro.value = structure.introducao || "";
  fields.application.value = structure.aplicacao || "";
  fields.challenge.value = structure.desafio || "";
  fields.appeal.value = structure.apelo || "";
  fields.titleSize.value = structure.estilo?.titulo || 54;
  fields.slideTitleSize.value = structure.estilo?.titulo_slide || 34;
  fields.bodySize.value = structure.estilo?.corpo || 22;
  renderCards();
}

function renderCards() {
  const holder = $("cardsEditor");
  holder.innerHTML = "";
  (structure.cards || []).forEach((card, index) => {
    const form = document.createElement("section");
    form.className = "card-form";
    form.innerHTML = `
      <div class="card-top">
        <label class="field">
          <span>Número</span>
          <input data-card="${index}" data-key="numero" value="${escapeAttr(card.numero)}" />
        </label>
        <label class="field">
          <span>Título</span>
          <input data-card="${index}" data-key="titulo" value="${escapeAttr(card.titulo)}" />
        </label>
        <label class="field">
          <span>Referência</span>
          <input data-card="${index}" data-key="referencia" value="${escapeAttr(card.referencia)}" />
        </label>
      </div>
      <label class="field">
        <span>Resumo</span>
        <textarea class="short" data-card="${index}" data-key="resumo">${escapeHtml(card.resumo)}</textarea>
      </label>
      <label class="field">
        <span>Frase de impacto</span>
        <input data-card="${index}" data-key="frase" value="${escapeAttr(card.frase)}" />
      </label>
    `;
    holder.appendChild(form);
  });
}

function collectStructure() {
  if (!structure) return null;
  structure.titulo = fields.title.value;
  structure.subtitulo = fields.subtitle.value;
  structure.texto_biblico = {
    referencia: fields.ref.value,
    texto: fields.verse.value,
  };
  structure.introducao = fields.intro.value;
  structure.aplicacao = fields.application.value;
  structure.desafio = fields.challenge.value;
  structure.apelo = fields.appeal.value;
  structure.estilo = {
    titulo: Number(fields.titleSize.value || 54),
    titulo_slide: Number(fields.slideTitleSize.value || 34),
    corpo: Number(fields.bodySize.value || 22),
  };

  document.querySelectorAll("[data-card]").forEach((input) => {
    const index = Number(input.dataset.card);
    const key = input.dataset.key;
    if (key === "numero") {
      structure.cards[index][key] = Number(input.value || index + 1);
    } else {
      structure.cards[index][key] = input.value;
    }
  });
  return structure;
}

function addCard() {
  if (!structure) return;
  structure.cards.push({
    numero: structure.cards.length + 1,
    titulo: "Novo card",
    referencia: "",
    resumo: "",
    frase: "",
  });
  renderCards();
}

async function exportFiles() {
  const payload = collectStructure();
  if (!payload) return;
  const coverImage = $("coverImageInput").files[0];
  payload.imagem_capa = coverImage ? await compressImage(coverImage) : (payload.imagem_capa || "");
  setStatus("Gerando arquivos...");
  $("exportBtn").disabled = true;
  clearGeneratedUrls();
  try {
    const baseName = slugify(payload.titulo || "sermao");
    const infographic = await requestFile("/api/export?kind=infografico", payload, `${baseName}_infografico.png`);
    const pdf = await requestFile("/api/export?kind=pdf", payload, `${baseName}_slides_cards.pdf`);
    const pptx = await requestFile("/api/pptx", payload, `${baseName}_slides_editaveis.pptx`);
    const projectBlob = new Blob(
      [JSON.stringify({ versao: 1, estrutura: payload }, null, 2)],
      { type: "application/json" },
    );
    const project = makeDownload(projectBlob, `${baseName}_projeto.json`);
    showDownloads({ infographic, pdf, pptx, project });
    setStatus("Arquivos gerados");
  } catch (error) {
    console.error(error);
    setStatus("Não foi possível gerar os arquivos");
  } finally {
    $("exportBtn").disabled = false;
  }
}

function showDownloads(assets) {
  $("preview").innerHTML = `<img src="${assets.infographic.url}" alt="Prévia do infográfico" />`;
  $("downloads").innerHTML = `
    <a class="download-link" href="${assets.infographic.url}" download="${assets.infographic.name}">Baixar infográfico PNG</a>
    <a class="download-link" href="${assets.pdf.url}" download="${assets.pdf.name}">Baixar PDF dos slides</a>
    <a class="download-link featured" href="${assets.pptx.url}" download="${assets.pptx.name}">Baixar PowerPoint editável</a>
    <a class="download-link project" href="${assets.project.url}" download="${assets.project.name}">Baixar projeto para reutilizar</a>
  `;
}

async function requestFile(endpoint, payload, name) {
  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) throw new Error(await response.text());
  return makeDownload(await response.blob(), name);
}

function makeDownload(blob, name) {
  const url = URL.createObjectURL(blob);
  generatedUrls.push(url);
  return { url, name };
}

function clearGeneratedUrls() {
  generatedUrls.forEach((url) => URL.revokeObjectURL(url));
  generatedUrls = [];
}

async function loadProject() {
  const file = $("projectInput").files[0];
  if (!file) return;
  try {
    const saved = JSON.parse(await file.text());
    structure = saved.estrutura || saved;
    populateEditor();
    $("exportBtn").disabled = false;
    $("addCardBtn").disabled = false;
    $("clearBtn").disabled = false;
    setStatus("Projeto restaurado");
  } catch {
    setStatus("Arquivo de projeto inválido");
  }
}

function clearWorkspace() {
  clearGeneratedUrls();
  structure = null;
  Object.values(fields).forEach((field) => {
    field.value = field.type === "number" ? field.defaultValue : "";
  });
  $("sermonText").value = "";
  $("docxInput").value = "";
  $("projectInput").value = "";
  $("coverImageInput").value = "";
  $("cardsEditor").innerHTML = "";
  $("preview").innerHTML = "<span>Os arquivos gerados aparecerão aqui.</span>";
  $("downloads").innerHTML = "";
  $("exportBtn").disabled = true;
  $("addCardBtn").disabled = true;
  $("clearBtn").disabled = true;
  setStatus("Aguardando sermão");
}

function readAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function compressImage(file) {
  const dataUrl = await readAsDataUrl(file);
  const image = await new Promise((resolve, reject) => {
    const element = new Image();
    element.onload = () => resolve(element);
    element.onerror = reject;
    element.src = dataUrl;
  });
  const scale = Math.min(1, 1600 / Math.max(image.width, image.height));
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(image.width * scale));
  canvas.height = Math.max(1, Math.round(image.height * scale));
  canvas.getContext("2d").drawImage(image, 0, 0, canvas.width, canvas.height);
  return canvas.toDataURL("image/jpeg", 0.82);
}

function slugify(value) {
  return String(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 80) || "sermao";
}

function setStatus(text) {
  $("status").textContent = text;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function escapeAttr(value) {
  return escapeHtml(value).replaceAll('"', "&quot;");
}
