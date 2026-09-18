import "./style.css";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import { prepareTemplateXml } from "./optional-fields.js";

const form = document.querySelector("#ficha-form");
const button = document.querySelector("#generate-button");
const errorBox = document.querySelector("#form-error");
const toast = document.querySelector("#toast");
const preview = document.querySelector("#document-preview");
const paperStage = document.querySelector("#paper-stage");
const spouseToggle = document.querySelector("#tem-conjuge");
const spouseFields = document.querySelector("#conjuge-fields");
const whatsappButton = document.querySelector("#whatsapp-button");
let generatedDocuments = null;

function onlyNumbers(value) { return String(value || "").replace(/\D/g, ""); }
function formatCpf(value) {
  return onlyNumbers(value).slice(0, 11)
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d)/, "$1.$2")
    .replace(/(\d{3})(\d{1,2})$/, "$1-$2");
}
function formatPhone(value) {
  const digits = onlyNumbers(value).slice(0, 11);
  if (digits.length <= 10) return digits.replace(/(\d{2})(\d{0,4})(\d{0,4})/, (_, a,b,c) => [a && `(${a})`, b, c].filter(Boolean).join(c ? "-" : " "));
  return digits.replace(/(\d{2})(\d{5})(\d{4})/, "($1) $2-$3");
}
function dateForDocument(value) {
  if (!value) return "";
  const [y,m,d] = value.split("-");
  return y && m && d ? `${d}/${m}/${y}` : value;
}
function clean(value) { return typeof value === "string" ? value.trim() : value; }
function getValues() {
  const raw = Object.fromEntries(new FormData(form).entries());
  const values = Object.fromEntries(Object.entries(raw).map(([k,v]) => [k, clean(v)]));
  values.tem_conjuge = spouseToggle.checked;
  values.loc_nascimento = dateForDocument(values.loc_nascimento);
  values.conj_nascimento = dateForDocument(values.conj_nascimento);
  if (!values.tem_conjuge) {
    for (const key of Object.keys(values)) if (key.startsWith("conj_")) values[key] = "";
  }
  return values;
}
function syncSpouseSection() {
  const enabled = spouseToggle.checked;
  spouseFields.hidden = !enabled;
  spouseFields.querySelectorAll("input, textarea, select").forEach((el) => { el.disabled = !enabled; });
  spouseFields.querySelectorAll("[data-conjuge-required]").forEach((el) => { el.required = enabled; });
  document.querySelector("#preview-conjuge").hidden = !enabled;
}
function updatePreview() {
  const d = getValues();
  document.querySelectorAll("[data-preview]").forEach((node) => {
    const key = node.dataset.preview;
    const fallback = key === "ficha_numero" ? "_____" : "";
    node.textContent = d[key] || fallback;
  });
  document.querySelector("#preview-loc-docs .preview-rg").hidden = !d.loc_rg;
  const conjRg = document.querySelector(".preview-conj-rg");
  if (conjRg) conjRg.hidden = !d.conj_rg;
  document.querySelector("#preview-conjuge").hidden = !d.tem_conjuge;
}
function resizePreview() {
  const available = paperStage.clientWidth - 32;
  const scale = Math.min(1, available / 794);
  preview.style.setProperty("--paper-scale", scale);
  paperStage.style.height = `${1123 * scale + 32}px`;
}
function safeFilename(value) {
  return String(value).replace(/[\\/*?:"<>|]/g, "").replace(/\s+/g, " ").trim().slice(0, 180);
}
async function loadTemplateBytes() {
  const partNames = ["part00.txt", "part01.txt", "part02.txt", "part03.txt", "part04.txt", "part05.txt"];
  const parts = await Promise.all(
    partNames.map(async (name) => {
      const response = await fetch(`${import.meta.env.BASE_URL}template/${name}`, { cache: "no-store" });
      if (!response.ok) throw new Error(`Parte do modelo não encontrada: ${name}`);
      return (await response.text()).trim();
    }),
  );

  const base64 = parts.join("").replace(/\s+/g, "");
  let binary;
  try {
    binary = atob(base64);
  } catch {
    throw new Error("O modelo Word não pôde ser decodificado.");
  }

  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);

  if (bytes.length !== 31916) {
    throw new Error(`Modelo Word incompleto (${bytes.length} de 31916 bytes).`);
  }
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error("O modelo Word não é um arquivo DOCX válido.");
  }

  return bytes;
}

async function createDocx(data) {
  const template = await loadTemplateBytes();
  const zip = new PizZip(template);
  zip.file("word/document.xml", prepareTemplateXml(zip.file("word/document.xml").asText(), data));
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: "{{", end: "}}" },
    parser(tag) {
      const key = tag.trim();
      return { get: (scope) => scope[key] ?? "" };
    },
  });
  doc.render(data);
  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}
async function createPdf() {
  const currentScale = preview.style.getPropertyValue("--paper-scale");
  preview.style.setProperty("--paper-scale", "1");
  await new Promise((resolve) => requestAnimationFrame(resolve));
  let canvas;
  try {
    canvas = await html2canvas(preview, {
      scale: 2,
      backgroundColor: "#ffffff",
      useCORS: true,
      logging: false,
      width: 794,
      height: 1123,
    });
  } finally {
    preview.style.setProperty("--paper-scale", currentScale || "1");
  }
  const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4", compress: true });
  pdf.addImage(canvas.toDataURL("image/jpeg", .96), "JPEG", 0, 0, 210, 297, undefined, "FAST");
  return pdf.output("blob");
}
async function createBundle(docxBlob, pdfBlob, baseName) {
  const zip = new PizZip();
  zip.file(`${baseName}.docx`, await docxBlob.arrayBuffer());
  zip.file(`${baseName}.pdf`, await pdfBlob.arrayBuffer());
  return zip.generate({ type: "blob", compression: "DEFLATE" });
}

function validateForm() {
  errorBox.hidden = true;
  syncSpouseSection();
  if (form.checkValidity()) return true;
  form.reportValidity();
  errorBox.textContent = "Preencha todos os campos obrigatórios antes de gerar os documentos.";
  errorBox.hidden = false;
  return false;
}

async function buildDocuments() {
  if (generatedDocuments) return generatedDocuments;

  const data = getValues();
  const baseName = safeFilename(`FICHA CADASTRAL ${data.loc_nome || "LOCATARIO"}`);
  updatePreview();
  await document.fonts.ready;

  const [docxBlob, pdfBlob] = await Promise.all([createDocx(data), createPdf()]);
  const docxFile = new File(
    [docxBlob],
    `${baseName}.docx`,
    { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
  );
  const pdfFile = new File(
    [pdfBlob],
    `${baseName}.pdf`,
    { type: "application/pdf" },
  );

  generatedDocuments = { data, baseName, docxBlob, pdfBlob, docxFile, pdfFile };
  return generatedDocuments;
}

function supportsNativeFileShare(files) {
  if (typeof navigator.share !== "function") return false;
  if (typeof navigator.canShare !== "function") return true;
  try {
    return navigator.canShare({ files });
  } catch {
    return false;
  }
}
function showToast(message) {
  toast.textContent = message;
  toast.classList.add("visible");
  window.setTimeout(() => toast.classList.remove("visible"), 3600);
}

for (const id of ["loc_cpf", "conj_cpf"]) {
  document.querySelector(`#${id}`).addEventListener("input", (event) => { event.target.value = formatCpf(event.target.value); });
}
for (const id of ["loc_cel", "conj_cel"]) {
  document.querySelector(`#${id}`).addEventListener("input", (event) => { event.target.value = formatPhone(event.target.value); });
}
spouseToggle.addEventListener("change", () => { syncSpouseSection(); updatePreview(); resizePreview(); });
form.addEventListener("input", () => {
  generatedDocuments = null;
  updatePreview();
});
form.addEventListener("change", () => {
  generatedDocuments = null;
  updatePreview();
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!validateForm()) return;

  button.disabled = true;
  whatsappButton.disabled = true;
  button.querySelector(".button-text").textContent = "Preparando documentos...";
  try {
    const docs = await buildDocuments();
    const bundle = await createBundle(docs.docxBlob, docs.pdfBlob, docs.baseName);
    saveAs(bundle, `${docs.baseName}.zip`);
    showToast("WORD e PDF gerados com sucesso.");
  } catch (error) {
    console.error(error);
    errorBox.textContent = `Não foi possível gerar os documentos: ${error.message}`;
    errorBox.hidden = false;
  } finally {
    button.disabled = false;
    whatsappButton.disabled = false;
    button.querySelector(".button-text").textContent = "Baixar WORD e PDF";
  }
});

whatsappButton.addEventListener("click", async () => {
  if (!validateForm()) return;

  const originalText = "Compartilhar no WhatsApp";
  whatsappButton.disabled = true;
  button.disabled = true;
  whatsappButton.querySelector(".whatsapp-button-text").textContent = "Preparando arquivos...";

  let fallbackWindow = null;
  const nativeShareExists = typeof navigator.share === "function";
  if (!nativeShareExists) {
    fallbackWindow = window.open("about:blank", "_blank");
  }

  try {
    const docs = await buildDocuments();
    const files = [docs.pdfFile, docs.docxFile];
    const message = `Olá! Segue a ficha cadastral preenchida de ${docs.data.loc_nome}.`;

    if (supportsNativeFileShare(files)) {
      try {
        await navigator.share({
          title: "Ficha Cadastral - Casas Comigo",
          text: message,
          files,
        });
        showToast("Arquivos prontos para compartilhar.");
        if (fallbackWindow) fallbackWindow.close();
        return;
      } catch (error) {
        if (error?.name === "AbortError") {
          if (fallbackWindow) fallbackWindow.close();
          return;
        }
        console.warn("Compartilhamento nativo falhou; usando fallback.", error);
      }
    }

    const bundle = await createBundle(docs.docxBlob, docs.pdfBlob, docs.baseName);
    saveAs(bundle, `${docs.baseName}.zip`);

    const whatsappUrl = `https://wa.me/?text=${encodeURIComponent(
      `${message}\n\nO arquivo com o PDF e o Word acabou de ser baixado. É só anexá-lo nesta conversa.`,
    )}`;

    if (fallbackWindow && !fallbackWindow.closed) {
      fallbackWindow.location.href = whatsappUrl;
    } else {
      window.open(whatsappUrl, "_blank", "noopener,noreferrer");
    }
    showToast("ZIP baixado. Anexe-o no WhatsApp.");
  } catch (error) {
    if (fallbackWindow && !fallbackWindow.closed) fallbackWindow.close();
    console.error(error);
    errorBox.textContent = `Não foi possível preparar o compartilhamento: ${error.message}`;
    errorBox.hidden = false;
  } finally {
    whatsappButton.disabled = false;
    button.disabled = false;
    whatsappButton.querySelector(".whatsapp-button-text").textContent = originalText;
  }
});

new ResizeObserver(resizePreview).observe(paperStage);
syncSpouseSection();
updatePreview();
resizePreview();
