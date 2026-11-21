// rbt-office demo: the package's own Word and spreadsheet shells, opened on a
// sample or on a file of the visitor's. Nothing leaves the page.
import * as Host from "@rbt/office/word/WordHost.js";
import * as Shell from "@rbt/office/word/WordShell.js";
import * as Dom from "@rbt/office/word/FableUI/Dom.js";
import Word from "@rbt/office/word/Word.js";
import createSpreadsheet from "@rbt/office/spreadsheet/Spreadsheet.js";
import { createSpreadsheetShell } from "@rbt/office/spreadsheet/Toolbar.js";
import { faceUrl, loadFaces, pdfFonts } from "./fonts.js";

const samples = [
  { file: "annual-review.docx", name: "Annual review", kind: "Document · report with photo and tables" },
  { file: "magazine-feature.docx", name: "Magazine feature", kind: "Document · columns, footnotes, pull quote, table" },
  { file: "field-notes.docx", name: "Field notes", kind: "Document · magazine page, drop cap, columns" },
  { file: "services-agreement.docx", name: "Services agreement", kind: "Document · formal, with placeholders" },
  { file: "dashboard.xlsx", name: "Generation dashboard", kind: "Spreadsheet · charts, data bars, formulas" },
  { file: "budget.xlsx", name: "Operating budget", kind: "Spreadsheet · a plain grid of formulas" },
];
// Macro-enabled files: VBA run in the page. `run` is the macro the bar offers first.
const macroSamples = [
  { file: "snake.xlsm", name: "Snake", kind: "Spreadsheet · VBA game: OnTime, OnKey", run: "StartGame" },
  { file: "mandelbrot.xlsm", name: "Mandelbrot", kind: "Spreadsheet · VBA, one cell a pixel", run: "DrawMandelbrot" },
  { file: "sierpinski.xlsm", name: "Sierpinski pentagon", kind: "Spreadsheet · VBA drawing 7,776 shapes", run: "main" },
  { file: "rainbow.docm", name: "Rainbow", kind: "Document · VBA: every character a colour", run: "Rainbow" },
];
const allSamples = [...samples, ...macroSamples];

const $ = (id) => document.getElementById(id);
// The TrueType faces the PDF writers embed: fetched once a document is on screen, not before
// (File > Print waits for them if it comes first).
let pdfFontsStarted = false;
function pdfFontsSoon(install) {
  setTimeout(() => pdfFonts().then(install), pdfFontsStarted ? 0 : 3000);
  pdfFontsStarted = true;
}

// The Word shell and the spreadsheet shell are each made once, on first use.
let word = null;
let sheet = null;
async function wordHost() {
  if (!word) {
    word = Host.create("", "word-shell");
    Dom.mountApp("#word", Shell.view(word));
    // File > Print renders a PDF with these fonts embedded.
    pdfFontsSoon((fonts) => Host.setPdfFonts(word, fonts));
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  }
  return word;
}
function sheetShell() {
  if (!sheet) {
    const grid = createSpreadsheet($("sheetgrid"));
    grid.Mount();
    // Faces the visitor's system lacks come from the page's free stand-ins.
    grid.SetFontLoader?.(async (family, bold, italic) => {
      const url = faceUrl(family, bold, italic);
      return url ? new Uint8Array(await (await fetch(url)).arrayBuffer()) : null;
    });
    pdfFontsSoon((fonts) => grid.SetPdfFonts(fonts));
    const shell = createSpreadsheetShell(grid, () => {});
    shell.Mount($("sheetbar"));
    sheet = { grid, shell };
  }
  return sheet;
}

function show(kind, name) {
  $("home").hidden = true;
  $("app").hidden = false;
  $("word").hidden = kind !== "word";
  $("sheet").hidden = kind !== "sheet";
  $("appname").textContent = name;
  document.title = `${name} · rbt-office`;
}

async function open(file, trusted = false) {
  const lower = file.name.toLowerCase();
  if (lower.endsWith(".docx") || lower.endsWith(".docm")) {
    show("word", file.name);
    // Only the faces this document is set in, loaded before it is laid out (its lines are measured
    // in them): read once here to ask the document which.
    const bytes = await file.arrayBuffer();
    const { forEditorPackage } = await import("@rbt/office/office-formats/src/WordPackageOpen.js");
    const [host, faces] = await Promise.all([wordHost(), forEditorPackage(bytes).then((p) => Word.getRequiredFonts(p)).catch(() => [])]);
    await loadFaces(faces);
    macroBar(null);
    Host.openFile(host, file);
    // (It opens asynchronously: the bar follows once the document is in.)
    for (let i = 0; i < 100 && /^Opening/.test(host.Status.Value); i++) await new Promise((r) => setTimeout(r, 100));
    macroBar(Word.hasMacros(host.Editor) ? wordMacros(host.Editor) : null, file.name, trusted);
  } else if (lower.endsWith(".xlsx") || lower.endsWith(".xlsm")) {
    show("sheet", file.name);
    const { grid, shell } = sheetShell();
    await grid.ImportXlsx(new Uint8Array(await file.arrayBuffer()));
    // (The shell names the workbook it saves after the file it opened itself.)
    shell.fileName = file.name;
    shell.refreshFileState?.();
    macroBar(grid.HasMacros ? sheetMacros(grid) : null, file.name, trusted);
  } else {
    alert("rbt-office opens .docx, .docm, .xlsx and .xlsm files.");
    return;
  }
  if (location.hash !== `#${encodeURIComponent(file.name)}`) history.pushState({ open: file.name }, "", `#${encodeURIComponent(file.name)}`);
}

// The macro bar: a file with macros says so, as Office's "Enable content"; once enabled, its macros
// (Alt+F8's list) run from here - or from the file's own buttons.
const sheetMacros = (grid) => ({
  enable: () => grid.EnableMacros(),
  list: () => grid.GetMacros(),
  run: (name) => grid.RunMacro(name),
  stop: () => grid.StopMacro(),
  running: () => grid.MacroRunning,
  focus: () => $("sheetgrid").querySelector("[tabindex]")?.focus(),
});
const wordMacros = (editor) => ({
  enable: () => Word.enableMacros(editor),
  list: () => Word.getMacros(editor),
  run: (name) => Word.runMacro(editor, name),
  stop: null,
  running: () => false,
  focus: () => {},
});
let macros = null;
// (The page's own samples open with their macros enabled; a visitor's file asks first, as Office does.)
function macroBar(m, fileName, trusted = false) {
  macros = m;
  const bar = $("macrobar");
  bar.hidden = !m;
  if (!m) return;
  $("macro-enable").hidden = trusted;
  $("macro-note").textContent = trusted ? "Enabling macros…" : "This file has VBA macros.";
  for (const id of ["macro-list", "macro-run", "macro-stop"]) $(id).hidden = true;
  bar.dataset.first = allSamples.find((s) => s.file === fileName)?.run ?? "";
  if (trusted) enableMacros();
}
$("macro-enable").addEventListener("click", () => enableMacros());
async function enableMacros() {
  const m = macros;
  if (!m) return;
  $("macro-enable").disabled = true;
  try { await m.enable(); } finally { $("macro-enable").disabled = false; }
  if (m !== macros) return;
  const list = $("macro-list");
  list.replaceChildren(...m.list().filter((n) => !/\.(Workbook|Document)_/.test(n)).map((n) => new Option(n, n)));
  const first = $("macrobar").dataset.first;
  if (first && [...list.options].some((o) => o.value === first)) list.value = first;
  $("macro-enable").hidden = true;
  $("macro-note").textContent = "Macros enabled.";
  list.hidden = $("macro-run").hidden = false;
  $("macro-stop").hidden = !m.stop;
  m.focus();
}
$("macro-run").addEventListener("click", () => { if (macros && $("macro-list").value) { macros.run($("macro-list").value); macros.focus(); } });
$("macro-stop").addEventListener("click", () => macros?.stop?.());

async function openSample(s) {
  const bytes = await (await fetch(`samples/${s.file}`)).arrayBuffer();
  await open(new File([bytes], s.file), true);
}

function home() {
  $("app").hidden = true;
  $("home").hidden = false;
  document.title = "rbt-office: documents and spreadsheets in the browser";
}

$("back").addEventListener("click", () => history.back());
addEventListener("popstate", () => (location.hash ? openLinked() : home()));

for (const [listId, group] of [["samples", samples], ["macro-samples", macroSamples]]) {
const list = $(listId);
for (const s of group) {
  const li = document.createElement("li");
  li.innerHTML = `<button type="button"><img alt="" src="samples/thumbs/${s.file}.png" loading="lazy"><span class="meta"><span class="name"></span><span class="kind"></span></span></button>`;
  li.querySelector(".name").textContent = s.name;
  li.querySelector(".kind").textContent = s.kind;
  li.querySelector("button").addEventListener("click", () => openSample(s));
  list.append(li);
}
}

$("pick").addEventListener("change", (e) => {
  const f = e.target.files?.[0];
  e.target.value = "";
  if (f) open(f);
});
const drop = $("drop");
drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  drop.classList.remove("over");
  const f = e.dataTransfer.files?.[0];
  if (f) open(f);
});

// A link to a sample (#annual-review.docx) opens it, on load or when the hash changes.
function openLinked() {
  const linked = allSamples.find((s) => `#${encodeURIComponent(s.file)}` === location.hash);
  // (The home page stays hidden meanwhile: its thumbnails are not fetched for a sample opened by link.)
  if (linked && $("appname").textContent !== linked.file) { $("home").hidden = true; openSample(linked); }
}
addEventListener("hashchange", openLinked);
openLinked();
window.demo = { open, openSample, get word() { return word; }, get sheet() { return sheet; } };
