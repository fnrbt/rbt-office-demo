// rbt-office demo: the package's own Word and spreadsheet shells, opened on a
// sample or on a file of the visitor's. Nothing leaves the page.
import "@rbt/office/office-formats/src/Main.js";
import * as Host from "@rbt/office/word/WordHost.js";
import * as Shell from "@rbt/office/word/WordShell.js";
import * as Dom from "@rbt/office/word/FableUI/Dom.js";
import createSpreadsheet from "@rbt/office/spreadsheet/Spreadsheet.js";
import { createSpreadsheetShell } from "@rbt/office/spreadsheet/Toolbar.js";
import { faceUrl, loadAllFaces } from "./fonts.js";

const samples = [
  { file: "annual-review.docx", name: "Annual review", kind: "Document · report with photo and tables" },
  { file: "field-notes.docx", name: "Field notes", kind: "Document · magazine page, drop cap, columns" },
  { file: "services-agreement.docx", name: "Services agreement", kind: "Document · formal, with placeholders" },
  { file: "dashboard.xlsx", name: "Generation dashboard", kind: "Spreadsheet · charts, data bars, formulas" },
  { file: "budget.xlsx", name: "Operating budget", kind: "Spreadsheet · a plain grid of formulas" },
];

const $ = (id) => document.getElementById(id);
const fontsReady = loadAllFaces();

// The Word shell and the spreadsheet shell are each made once, on first use.
let word = null;
let sheet = null;
async function wordHost() {
  if (!word) {
    word = Host.create("", "word-shell");
    Dom.mountApp("#word", Shell.view(word));
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

async function open(file) {
  const lower = file.name.toLowerCase();
  await fontsReady;
  if (lower.endsWith(".docx")) {
    show("word", file.name);
    Host.openFile(await wordHost(), file);
  } else if (lower.endsWith(".xlsx")) {
    show("sheet", file.name);
    const { grid, shell } = sheetShell();
    await grid.ImportXlsx(new Uint8Array(await file.arrayBuffer()));
    // (The shell names the workbook it saves after the file it opened itself.)
    shell.fileName = file.name;
    shell.refreshFileState?.();
  } else {
    alert("rbt-office opens .docx and .xlsx files.");
    return;
  }
  if (location.hash !== `#${encodeURIComponent(file.name)}`) history.pushState({ open: file.name }, "", `#${encodeURIComponent(file.name)}`);
}

async function openSample(s) {
  const bytes = await (await fetch(`samples/${s.file}`)).arrayBuffer();
  await open(new File([bytes], s.file));
}

function home() {
  $("app").hidden = true;
  $("home").hidden = false;
  document.title = "rbt-office: documents and spreadsheets in the browser";
}

$("back").addEventListener("click", () => history.back());
addEventListener("popstate", () => (location.hash ? openLinked() : home()));

const list = $("samples");
for (const s of samples) {
  const li = document.createElement("li");
  li.innerHTML = `<button type="button"><img alt="" src="samples/thumbs/${s.file}.png" loading="lazy"><span class="meta"><span class="name"></span><span class="kind"></span></span></button>`;
  li.querySelector(".name").textContent = s.name;
  li.querySelector(".kind").textContent = s.kind;
  li.querySelector("button").addEventListener("click", () => openSample(s));
  list.append(li);
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
  const linked = samples.find((s) => `#${encodeURIComponent(s.file)}` === location.hash);
  if (linked && $("appname").textContent !== linked.file) openSample(linked);
}
addEventListener("hashchange", openLinked);
openLinked();
window.demo = { open, openSample, get word() { return word; }, get sheet() { return sheet; } };
