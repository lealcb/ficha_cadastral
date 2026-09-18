import "./style.css";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import { omitOptionalRg } from "./optional-fields.js";

const form = document.querySelector("#ficha-form");
const button = document.querySelector("#generate-button");
const errorBox = document.querySelector("#form-error");
const toast = document.querySelector("#toast");
const preview = document.querySelector("#document-preview");
const paperStage = document.querySelector("#paper-stage");
const spouseToggle = document.querySelector("#tem-conjuge");
const spouseFields = document.querySelector("#conjuge-fields");

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
async function createDocx(data) {
  const response = await fetch(`${import.meta.env.BASE_URL}TEMPLATE.docx`);
  if (!response.ok) throw new Error("O modelo DOCX não foi encontrado.");
  const template = await response.arrayBuffer();
  const zip = new PizZip(template);
  zip.file("word/document.xml", omitOptionalRg(zip.file("word/document.xml").asText(), data));
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
  const logo = preview.querySelector("img");
  if (logo?.decode) await logo.decode();
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
form.addEventListener("input", updatePreview);
form.addEventListener("change", updatePreview);

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  errorBox.hidden = true;
  syncSpouseSection();
  if (!form.checkValidity()) {
    form.reportValidity();
    errorBox.textContent = "Preencha todos os campos obrigatórios antes de gerar os documentos.";
    errorBox.hidden = false;
    return;
  }

  const data = getValues();
  const baseName = safeFilename(`FICHA CADASTRAL ${data.loc_nome || "LOCATARIO"}`);
  button.disabled = true;
  button.querySelector(".button-text").textContent = "Preparando documentos...";
  try {
    updatePreview();
    await document.fonts.ready;
    const [docx, pdf] = await Promise.all([createDocx(data), createPdf()]);
    const bundle = await createBundle(docx, pdf, baseName);
    saveAs(bundle, `${baseName}.zip`);
    showToast("WORD e PDF gerados com sucesso.");
  } catch (error) {
    console.error(error);
    errorBox.textContent = `Não foi possível gerar os documentos: ${error.message}`;
    errorBox.hidden = false;
  } finally {
    button.disabled = false;
    button.querySelector(".button-text").textContent = "Baixar WORD e PDF";
  }
});

new ResizeObserver(resizePreview).observe(paperStage);
syncSpouseSection();
updatePreview();
resizePreview();
