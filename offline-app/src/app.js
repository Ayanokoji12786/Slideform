import * as pdfjsLib from "./vendor/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.mjs";
if ("serviceWorker" in navigator) navigator.serviceWorker.register("./service-worker.js");
const ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
const state = { file: null, kind: null, slides: [], tableMode: "first", pdf: null };
const $ = (selector) => document.querySelector(selector);
const fileInput = $("#pptx-file");
const dropzone = $("#dropzone");
const fileName = $("#file-name");
const slideNumbers = $("#slide-numbers");
const convertButton = $("#convert-button");
const status = $("#status");
const selectionHelp = $("#selection-help");
const choices = $(".choice-row");

const normalise = (value) => value.replace(/\s+/g, " ").trim().toLocaleLowerCase();
const nodes = (node, name) => Array.from(node.getElementsByTagNameNS(ns, name));
const slideNumber = (path) => Number(path.match(/slide(\d+)\.xml$/)[1]);
const outputName = () => state.file.name.replace(/\.(pptx|pdf)$/i, ".xlsx");

function setStatus(message, error = false) {
  status.textContent = message;
  status.classList.toggle("error", error);
}

function textBodyText(body) {
  return nodes(body, "p").map((paragraph) => nodes(paragraph, "t").map((run) => run.textContent).join("")).join("\n");
}

function downloadBuffer(buffer, name) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  link.download = name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

async function readPptx(file) {
  const zip = await JSZip.loadAsync(file);
  const paths = Object.keys(zip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path)).sort((a, b) => slideNumber(a) - slideNumber(b));
  const parser = new DOMParser();
  const slides = [];
  for (const path of paths) {
    const xml = parser.parseFromString(await zip.file(path).async("text"), "application/xml");
    if (xml.querySelector("parsererror")) throw new Error("One slide could not be read.");
    const tables = nodes(xml, "tbl").map((table) => nodes(table, "tr").map((row) => nodes(row, "tc").map(textBodyText)));
    const content = [
      ...nodes(xml, "txBody").map(textBodyText),
      ...tables.flatMap((table) => table.map((row) => row.join(" "))),
    ];
    slides.push({ number: slideNumber(path), tables, hasDemandPriority: content.some((value) => normalise(value).includes("demand priority")) });
  }
  return slides;
}

async function readPdf(file) {
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  return { pdf, pages: Array.from({ length: pdf.numPages }, (_, index) => index + 1) };
}

function selectedNumbers() {
  const values = slideNumbers.value.split(",").map((item) => item.trim()).filter(Boolean);
  if (!values.length || values.some((item) => !/^\d+$/.test(item) || Number(item) < 1)) throw new Error("Enter positive slide numbers separated by commas, for example: 6, 11, 15.");
  return [...new Set(values.map(Number))];
}

async function setFile(file) {
  const pptx = file?.name.toLowerCase().endsWith(".pptx");
  const pdf = file?.name.toLowerCase().endsWith(".pdf");
  if (!pptx && !pdf) return setStatus("Choose a .pptx PowerPoint file or a rendered PDF.", true);
  state.file = file; state.kind = pdf ? "pdf" : "pptx"; state.slides = []; state.pdf = null;
  fileName.textContent = file.name; convertButton.disabled = true;
  setStatus(pdf ? "Reading PDF pages…" : "Reading native slide content…");
  try {
    if (pdf) {
      const result = await readPdf(file);
      state.pdf = result.pdf; slideNumbers.value = result.pages.join(", "); choices.hidden = true;
      selectionHelp.textContent = "This PDF was rendered by a presentation app. Select the pages to place in Excel as slide images.";
      setStatus(`Selected all ${result.pages.length} PDF pages for image export.`);
    } else {
      state.slides = await readPptx(file);
      const matches = state.slides.filter((slide) => slide.hasDemandPriority).map((slide) => slide.number);
      slideNumbers.value = matches.join(", "); choices.hidden = false;
      selectionHelp.innerHTML = 'Slides containing <strong>Demand Priority</strong> are selected automatically. Edit the list if needed.';
      setStatus(matches.length ? `Selected slides ${matches.join(", ")} because they contain Demand Priority.` : "No Demand Priority slides found. Enter slide numbers manually.");
    }
    convertButton.disabled = false;
  } catch (error) {
    state.file = null; state.slides = []; state.pdf = null; fileName.textContent = "No file selected";
    setStatus(error.message || "This file could not be read.", true);
  }
}

function convertTables(selected) {
  const slides = selected.map((number) => state.slides.find((slide) => slide.number === number));
  if (slides.some((slide) => !slide)) throw new Error("One or more selected slide numbers do not exist in this presentation.");
  const workbook = XLSX.utils.book_new(); let count = 0;
  slides.forEach((slide) => {
    const tables = state.tableMode === "all" ? slide.tables : slide.tables.slice(0, 1);
    if (!tables.length) throw new Error(`Slide ${slide.number} does not contain a native PowerPoint table.`);
    tables.forEach((rows, index) => {
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), `Slide ${slide.number} Table ${index + 1}`);
      count += 1;
    });
  });
  XLSX.writeFile(workbook, outputName(), { compression: true });
  return `${count} native table${count === 1 ? "" : "s"}`;
}

async function convertPdfPages(selected) {
  if (selected.some((number) => number > state.pdf.numPages)) throw new Error("One or more selected page numbers do not exist in this PDF.");
  const workbook = new ExcelJS.Workbook();
  for (const number of selected) {
    setStatus(`Rendering PDF page ${number}…`);
    const page = await state.pdf.getPage(number);
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
    await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
    const sheet = workbook.addWorksheet(`Slide ${number}`, { views: [{ showGridLines: false }] });
    sheet.getColumn(1).width = 135;
    const image = workbook.addImage({ base64: canvas.toDataURL("image/png"), extension: "png" });
    const width = 960;
    sheet.addImage(image, { tl: { col: 0, row: 0 }, ext: { width, height: Math.round(width * canvas.height / canvas.width) } });
  }
  downloadBuffer(await workbook.xlsx.writeBuffer(), outputName());
  return `${selected.length} rendered slide image${selected.length === 1 ? "" : "s"}`;
}

async function convert() {
  try {
    const selected = selectedNumbers();
    convertButton.disabled = true; setStatus(`Creating ${outputName()}…`);
    const result = state.kind === "pdf" ? await convertPdfPages(selected) : convertTables(selected);
    setStatus(`Downloaded ${outputName()} with ${result}.`);
  } catch (error) {
    setStatus(error.message || "The workbook could not be created.", true);
  } finally {
    convertButton.disabled = !state.file;
  }
}

fileInput.addEventListener("change", () => setFile(fileInput.files[0]));
["dragenter", "dragover"].forEach((event) => dropzone.addEventListener(event, (e) => { e.preventDefault(); dropzone.classList.add("dragging"); }));
["dragleave", "drop"].forEach((event) => dropzone.addEventListener(event, (e) => { e.preventDefault(); dropzone.classList.remove("dragging"); }));
dropzone.addEventListener("drop", (event) => setFile(event.dataTransfer.files[0]));
document.querySelectorAll(".choice").forEach((button) => button.addEventListener("click", () => {
  state.tableMode = button.dataset.tableMode;
  document.querySelectorAll(".choice").forEach((choice) => {
    const active = choice === button; choice.classList.toggle("active", active); choice.setAttribute("aria-pressed", String(active));
  });
}));
convertButton.addEventListener("click", convert);
