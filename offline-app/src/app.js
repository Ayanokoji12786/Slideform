import * as pdfjsLib from "./vendor/pdf.min.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = "./vendor/pdf.worker.min.mjs";
if ("serviceWorker" in navigator) navigator.serviceWorker.register("./service-worker.js");
const ns = "http://schemas.openxmlformats.org/drawingml/2006/main";
const presentationNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
const chartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const diagramNs = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const state = { file: null, kind: null, slides: [], tableMode: "all", pdf: null };
const $ = (selector) => document.querySelector(selector);
const fileInput = $("#pptx-file");
const dropzone = $("#dropzone");
const fileName = $("#file-name");
const slideNumbers = $("#slide-numbers");
const convertButton = $("#convert-button");
const status = $("#status");
const selectionHelp = $("#selection-help");
const visualNote = $("#visual-copy-note");
const choices = $(".choice-row");

const normalise = (value) => value.replace(/\s+/g, " ").trim().toLocaleLowerCase();
const nodes = (node, name) => Array.from(node.getElementsByTagNameNS(ns, name));
const presentationNodes = (node, name) => Array.from(node.getElementsByTagNameNS(presentationNs, name));
const chartNodes = (node, name) => Array.from(node.getElementsByTagNameNS(chartNs, name));
const slideNumber = (path) => Number(path.match(/slide(\d+)\.xml$/)[1]);
const outputName = () => state.file.name.replace(/\.(pptx|pdf)$/i, ".xlsx");

function setStatus(message, error = false) {
  status.textContent = message;
  status.classList.toggle("error", error);
}

function textBodyText(body) {
  return nodes(body, "p").map((paragraph) => nodes(paragraph, "t").map((run) => run.textContent).join("")).join("\n");
}

function geometry(node) {
  const transform = presentationNodes(node, "xfrm")[0] || nodes(node, "xfrm")[0];
  const offset = transform && nodes(transform, "off")[0];
  const extent = transform && nodes(transform, "ext")[0];
  return {
    x: Number(offset?.getAttribute("x") || 0),
    y: Number(offset?.getAttribute("y") || 0),
    width: Number(extent?.getAttribute("cx") || 0),
    height: Number(extent?.getAttribute("cy") || 0),
  };
}

function resolvePptPath(base, target) {
  return new URL(target, `https://slideform.local/${base}`).pathname.slice(1);
}

async function readSlideNotes(zip, number, parser) {
  const relPath = `ppt/slides/_rels/slide${number}.xml.rels`;
  if (!zip.file(relPath)) return "";
  const rels = parser.parseFromString(await zip.file(relPath).async("text"), "application/xml");
  const relationship = Array.from(rels.getElementsByTagName("Relationship")).find((r) => (r.getAttribute("Type") || "").endsWith("/notesSlide"));
  if (!relationship) return "";
  const notesFile = zip.file(resolvePptPath("ppt/slides/", relationship.getAttribute("Target")));
  if (!notesFile) return "";
  const notesXml = parser.parseFromString(await notesFile.async("text"), "application/xml");
  const bodies = presentationNodes(notesXml, "sp").map((shape) => {
    const placeholderType = presentationNodes(shape, "ph")[0]?.getAttribute("type");
    if (placeholderType === "sldNum" || placeholderType === "sldImg") return "";
    const text = presentationNodes(shape, "txBody")[0] || nodes(shape, "txBody")[0];
    return text ? textBodyText(text) : "";
  }).filter((value) => value.trim());
  return bodies.join("\n").trim();
}

function cachedPoints(container) {
  if (!container) return [];
  const cache = chartNodes(container, "strCache")[0] || chartNodes(container, "numCache")[0];
  if (!cache) return [];
  return chartNodes(cache, "pt").sort((a, b) => Number(a.getAttribute("idx")) - Number(b.getAttribute("idx"))).map((point) => chartNodes(point, "v")[0]?.textContent || "");
}

async function readSlideCharts(zip, number, xml, parser) {
  const relPath = `ppt/slides/_rels/slide${number}.xml.rels`;
  if (!zip.file(relPath)) return [];
  const rels = parser.parseFromString(await zip.file(relPath).async("text"), "application/xml");
  const targets = new Map(Array.from(rels.getElementsByTagName("Relationship")).map((relationship) => [relationship.getAttribute("Id"), relationship.getAttribute("Target")]));
  const charts = [];
  for (const frame of presentationNodes(xml, "graphicFrame")) {
    const chartRef = Array.from(frame.getElementsByTagNameNS(chartNs, "chart"))[0];
    const relationId = chartRef?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    const target = relationId && targets.get(relationId);
    if (!target) continue;
    const chartFile = zip.file(resolvePptPath("ppt/slides/", target));
    if (!chartFile) continue;
    const chartXml = parser.parseFromString(await chartFile.async("text"), "application/xml");
    const rows = [];
    for (const series of Array.from(chartXml.getElementsByTagNameNS(chartNs, "ser"))) {
      const name = Array.from(series.getElementsByTagNameNS(chartNs, "v"))[0]?.textContent || "";
      const categories = cachedPoints(Array.from(series.getElementsByTagNameNS(chartNs, "cat"))[0]);
      const values = cachedPoints(Array.from(series.getElementsByTagNameNS(chartNs, "val"))[0]);
      const length = Math.max(categories.length, values.length, 1);
      for (let index = 0; index < length; index += 1) rows.push([name, categories[index] || "", values[index] || ""]);
    }
    if (rows.length) charts.push({ geometry: geometry(frame), rows });
  }
  return charts;
}

const LABEL_LINE_RE = /^[A-Za-z][A-Za-z0-9 &/'()-]{1,55}:$/;
const LABEL_INLINE_RE = /(?:^|\n)([A-Za-z][A-Za-z0-9 &/'()-]{1,55}):[ \t]*\n?/g;

function collapse(text) {
  return (text || "").replace(/\s+/g, " ").trim();
}

function slideTitle(xml) {
  const shape = presentationNodes(xml, "sp")[0];
  const text = shape && (presentationNodes(shape, "txBody")[0] || nodes(shape, "txBody")[0]);
  return text ? collapse(textBodyText(text)) : "";
}

function shapeAnchor(g) {
  return { x: g.x + g.width / 2, y: g.y + g.height / 2 };
}

function extractStructuredFields(xml) {
  const fields = [];
  const titleShape = presentationNodes(xml, "sp")[0];
  const claimed = new Set();

  const freeShapes = presentationNodes(xml, "sp").filter((shape) => shape !== titleShape).map((shape) => {
    const text = presentationNodes(shape, "txBody")[0] || nodes(shape, "txBody")[0];
    const value = text && textBodyText(text);
    return value && value.trim() ? { shape, geometry: geometry(shape), text: value } : null;
  }).filter(Boolean);

  presentationNodes(xml, "graphicFrame").forEach((frame) => {
    const table = nodes(frame, "tbl")[0];
    if (!table) return;
    const rows = nodes(table, "tr").map((row) => nodes(row, "tc").map(textBodyText));

    if (rows.length === 2 && rows[0].length === 1 && rows[1].length === 1) {
      // A 2-row / 1-column "header card": label in row 0, value either inline in row 1
      // or in the nearest other shape on the slide (common when a slide styles the label
      // and its value as two separate boxes with no structural link between them).
      const label = rows[0][0].trim();
      if (!label || label.length > 60) return;
      const inline = rows[1][0].trim();
      if (inline) { fields.push({ label, value: collapse(inline) }); return; }
      const anchor = shapeAnchor(geometry(frame));
      let best = null; let bestDistance = Infinity;
      freeShapes.forEach((candidate) => {
        if (claimed.has(candidate.shape) || LABEL_LINE_RE.test(candidate.text.trim())) return;
        const c = shapeAnchor(candidate.geometry);
        const distance = Math.hypot(c.x - anchor.x, c.y - anchor.y);
        if (distance < bestDistance) { bestDistance = distance; best = candidate; }
      });
      if (best) { claimed.add(best.shape); fields.push({ label, value: collapse(best.text) }); }
      return;
    }

    for (let i = 0; i < rows.length; i += 1) {
      const [first, second] = rows[i];
      if (!first || !LABEL_LINE_RE.test(first.trim())) continue;
      const label = first.trim().replace(/:$/, "");
      if (second && second.trim()) {
        fields.push({ label, value: collapse(second) });
      } else if (rows[i + 1]?.[0]?.trim() && !LABEL_LINE_RE.test(rows[i + 1][0].trim())) {
        fields.push({ label, value: collapse(rows[i + 1][0]) });
        i += 1;
      }
    }
  });

  freeShapes.forEach((candidate) => {
    if (claimed.has(candidate.shape)) return;
    const matches = Array.from(candidate.text.matchAll(LABEL_INLINE_RE));
    if (!matches.length) return;
    matches.forEach((match, index) => {
      const start = match.index + match[0].length;
      const end = index + 1 < matches.length ? matches[index + 1].index : candidate.text.length;
      const value = collapse(candidate.text.slice(start, end));
      if (value) fields.push({ label: match[1].trim(), value });
    });
  });

  return fields;
}

function visualOnlyReasons(xml, chartsFound) {
  const reasons = new Set();
  if (xml.getElementsByTagNameNS(diagramNs, "relIds").length) reasons.add("a SmartArt diagram");
  if (presentationNodes(xml, "cxnSp").some((shape) => nodes(shape, "stCxn").length || nodes(shape, "endCxn").length)) reasons.add("connected shapes");
  if (nodes(xml, "graphicData").some((data) => (data.getAttribute("uri") || "").toLowerCase().includes("ole"))) reasons.add("an embedded object");
  const chartFrames = presentationNodes(xml, "graphicFrame").filter((frame) => frame.getElementsByTagNameNS(chartNs, "chart").length).length;
  if (chartFrames > chartsFound) reasons.add("a chart without readable cached data");
  return Array.from(reasons);
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
  const presentationFile = zip.file("ppt/presentation.xml");
  if (!presentationFile) throw new Error("This file is not a valid PowerPoint presentation.");
  const parser = new DOMParser();
  const presentation = parser.parseFromString(await presentationFile.async("text"), "application/xml");
  const slideSize = presentationNodes(presentation, "sldSz")[0];
  const size = { width: Number(slideSize?.getAttribute("cx")) || 12192000, height: Number(slideSize?.getAttribute("cy")) || 6858000 };
  const slides = [];
  for (const path of paths) {
    const xml = parser.parseFromString(await zip.file(path).async("text"), "application/xml");
    if (xml.querySelector("parsererror")) throw new Error("One slide could not be read.");
    const number = slideNumber(path);
    const tables = presentationNodes(xml, "graphicFrame").map((frame) => {
      const table = nodes(frame, "tbl")[0];
      return table && { geometry: geometry(frame), rows: nodes(table, "tr").map((row) => nodes(row, "tc").map(textBodyText)) };
    }).filter(Boolean);
    const textBoxes = presentationNodes(xml, "sp").map((shape) => {
      const text = presentationNodes(shape, "txBody")[0] || nodes(shape, "txBody")[0];
      return text && { geometry: geometry(shape), text: textBodyText(text) };
    }).filter((shape) => shape && shape.text);
    const charts = await readSlideCharts(zip, number, xml, parser);
    const notes = await readSlideNotes(zip, number, parser);
    const content = [
      ...textBoxes.map((shape) => shape.text),
      ...tables.flatMap((table) => table.rows.map((row) => row.join(" "))),
    ];
    const visualOnly = visualOnlyReasons(xml, charts.length);
    const title = slideTitle(xml);
    const structuredFields = extractStructuredFields(xml);
    slides.push({ number, size, tables, textBoxes, charts, visualOnly, title, structuredFields, notes, hasDemandPriority: content.some((value) => normalise(value).includes("demand priority")) });
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

function updateVisualCopyNote() {
  if (state.kind !== "pptx" || !state.slides.length) { visualNote.hidden = true; return; }
  const numbers = slideNumbers.value.split(",").map((item) => item.trim()).filter((item) => /^\d+$/.test(item)).map(Number);
  const flagged = numbers.map((number) => state.slides.find((slide) => slide.number === number)).filter((slide) => slide?.visualOnly.length);
  if (!flagged.length) { visualNote.hidden = true; return; }
  const list = flagged.map((slide) => slide.number).join(", ");
  const pronoun = flagged.length === 1 ? "it" : "them";
  visualNote.textContent = `Slide${flagged.length === 1 ? "" : "s"} ${list} also ${flagged.length === 1 ? "has" : "have"} content native extraction can't fully capture. Export ${pronoun} to PDF separately for a faithful copy.`;
  visualNote.hidden = false;
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
      visualNote.hidden = true;
    } else {
      state.slides = await readPptx(file);
      const matches = state.slides.filter((slide) => slide.hasDemandPriority).map((slide) => slide.number);
      slideNumbers.value = matches.join(", "); choices.hidden = false;
      selectionHelp.innerHTML = 'Slides containing <strong>Demand Priority</strong> are selected automatically. Edit the list if needed.';
      setStatus(matches.length ? `Selected slides ${matches.join(", ")} because they contain Demand Priority.` : "No Demand Priority slides found. Enter slide numbers manually.");
      updateVisualCopyNote();
    }
    convertButton.disabled = false;
  } catch (error) {
    state.file = null; state.slides = []; state.pdf = null; fileName.textContent = "No file selected";
    setStatus(error.message || "This file could not be read.", true);
  }
}

function slidePosition(value, total, available) {
  return Math.max(1, Math.min(available, Math.floor((value / total) * available) + 1));
}

function reserveArea(occupied, row, column, rowSpan, columnSpan) {
  let targetRow = row;
  const free = () => {
    for (let r = targetRow; r < targetRow + rowSpan; r += 1) for (let c = column; c < column + columnSpan; c += 1) if (occupied.has(`${r}:${c}`)) return false;
    return true;
  };
  while (!free()) targetRow += 1;
  for (let r = targetRow; r < targetRow + rowSpan; r += 1) for (let c = column; c < column + columnSpan; c += 1) occupied.add(`${r}:${c}`);
  return targetRow;
}

function applyBorder(cell) {
  cell.border = { top: { style: "thin", color: { argb: "FFBFC3C9" } }, left: { style: "thin", color: { argb: "FFBFC3C9" } }, bottom: { style: "thin", color: { argb: "FFBFC3C9" } }, right: { style: "thin", color: { argb: "FFBFC3C9" } } };
  cell.alignment = { vertical: "top", wrapText: true };
}

async function convertStructuredSummary(selected) {
  const slides = selected.map((number) => state.slides.find((slide) => slide.number === number));
  if (slides.some((slide) => !slide)) throw new Error("One or more selected slide numbers do not exist in this presentation.");

  const columns = [];
  const addColumn = (name) => { if (!columns.includes(name)) columns.push(name); };
  addColumn("Slide");
  addColumn("Name");

  // A note repeated verbatim on more than one selected slide is almost always a leftover
  // from duplicating a slide as a template, not real per-slide content - surfacing it as
  // if it were specific to each slide would be actively misleading, so it's dropped.
  const noteCounts = new Map();
  slides.forEach((slide) => { if (slide.notes) noteCounts.set(slide.notes, (noteCounts.get(slide.notes) || 0) + 1); });

  const rows = slides.map((slide) => {
    const row = { Slide: slide.number, Name: slide.title };
    slide.structuredFields.forEach(({ label, value }) => { addColumn(label); row[label] = row[label] ? `${row[label]}\n${value}` : value; });
    if (slide.notes && noteCounts.get(slide.notes) === 1) { addColumn("Speaker Notes"); row["Speaker Notes"] = slide.notes; }
    return row;
  });

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Demands", { views: [{ showGridLines: false, state: "frozen", ySplit: 1 }] });
  columns.forEach((name, index) => { sheet.getColumn(index + 1).width = name === "Slide" ? 8 : 32; });
  const header = sheet.getRow(1);
  columns.forEach((name, index) => { const cell = header.getCell(index + 1); cell.value = name; cell.font = { bold: true, color: { argb: "FFFFFFFF" } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF111420" } }; applyBorder(cell); });
  rows.forEach((row, rowIndex) => {
    columns.forEach((name, columnIndex) => {
      const cell = sheet.getCell(rowIndex + 2, columnIndex + 1);
      cell.value = row[name] ?? "";
      applyBorder(cell);
    });
  });

  return workbook.xlsx.writeBuffer().then((buffer) => { downloadBuffer(buffer, outputName()); return `${rows.length} row${rows.length === 1 ? "" : "s"}`; });
}

async function convertPptxContent(selected) {
  const slides = selected.map((number) => state.slides.find((slide) => slide.number === number));
  if (slides.some((slide) => !slide)) throw new Error("One or more selected slide numbers do not exist in this presentation.");
  const workbook = new ExcelJS.Workbook(); let count = 0;
  slides.forEach((slide) => {
    const sheet = workbook.addWorksheet(`Slide ${slide.number}`, { views: [{ showGridLines: false }] });
    const columns = 18; const rows = 36; const occupied = new Set();
    for (let column = 1; column <= columns; column += 1) sheet.getColumn(column).width = 12;
    for (let row = 1; row <= rows; row += 1) sheet.getRow(row).height = 19;

    slide.textBoxes.forEach((shape) => {
      const column = slidePosition(shape.geometry.x, slide.size.width, columns);
      const row = slidePosition(shape.geometry.y, slide.size.height, rows);
      const columnSpan = Math.max(1, Math.ceil((shape.geometry.width / slide.size.width) * columns));
      const rowSpan = Math.max(1, Math.ceil((shape.geometry.height / slide.size.height) * rows));
      const targetRow = reserveArea(occupied, row, column, rowSpan, Math.min(columnSpan, columns - column + 1));
      const cell = sheet.getCell(targetRow, column);
      cell.value = shape.text;
      cell.alignment = { vertical: "top", wrapText: true };
    });

    const tables = state.tableMode === "all" ? slide.tables : slide.tables.slice(0, 1);
    tables.forEach((table) => {
      const column = slidePosition(table.geometry.x, slide.size.width, columns);
      const row = slidePosition(table.geometry.y, slide.size.height, rows);
      const columnSpan = Math.max(1, Math.ceil((table.geometry.width / slide.size.width) * columns));
      const tableColumns = Math.max(...table.rows.map((items) => items.length), 1);
      const targetRow = reserveArea(occupied, row, column, Math.max(table.rows.length, 1), Math.min(Math.max(columnSpan, tableColumns), columns - column + 1));
      table.rows.forEach((items, rowIndex) => items.forEach((value, columnIndex) => {
        const cell = sheet.getCell(targetRow + rowIndex, column + columnIndex);
        cell.value = value;
        applyBorder(cell);
      }));
    });
    slide.charts.forEach((chart) => {
      const column = slidePosition(chart.geometry.x, slide.size.width, columns);
      const row = slidePosition(chart.geometry.y, slide.size.height, rows);
      const targetRow = reserveArea(occupied, row, column, chart.rows.length + 1, 3);
      ["Series", "Category", "Value"].forEach((value, index) => { const cell = sheet.getCell(targetRow, column + index); cell.value = value; applyBorder(cell); });
      chart.rows.forEach((items, rowIndex) => items.forEach((value, columnIndex) => { const cell = sheet.getCell(targetRow + rowIndex + 1, column + columnIndex); cell.value = value; applyBorder(cell); }));
    });
    count += 1;
  });
  return workbook.xlsx.writeBuffer().then((buffer) => { downloadBuffer(buffer, outputName()); return `${count} slide worksheet${count === 1 ? "" : "s"}`; });
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
    const result = state.kind === "pdf" ? await convertPdfPages(selected) : state.tableMode === "structured" ? await convertStructuredSummary(selected) : await convertPptxContent(selected);
    setStatus(`Downloaded ${outputName()} with ${result}.`);
  } catch (error) {
    setStatus(error.message || "The workbook could not be created.", true);
  } finally {
    convertButton.disabled = !state.file;
  }
}

fileInput.addEventListener("change", () => setFile(fileInput.files[0]));
slideNumbers.addEventListener("input", updateVisualCopyNote);
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
