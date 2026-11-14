// Free, metrically compatible stand-ins for the fonts Office documents name,
// registered under those names, so documents lay out with Office's line breaks
// without shipping Microsoft's fonts. (Licences: public/fonts/*.txt.)
//   Calibri -> Carlito, Cambria -> Caladea, Georgia -> Gelasio,
//   Times New Roman -> Liberation Serif, Arial -> Liberation Sans, Segoe UI -> Selawik.
const R = "regular", B = "bold", I = "italic", BI = "boldItalic";
const families = {
  Calibri: { [R]: "Carlito-Regular", [B]: "Carlito-Bold", [I]: "Carlito-Italic", [BI]: "Carlito-BoldItalic" },
  "Calibri Light": { [R]: "Carlito-Regular", [B]: "Carlito-Bold", [I]: "Carlito-Italic", [BI]: "Carlito-BoldItalic" },
  Cambria: { [R]: "Caladea-Regular", [B]: "Caladea-Bold", [I]: "Caladea-Italic", [BI]: "Caladea-BoldItalic" },
  Georgia: { [R]: "Gelasio-Var", [B]: "Gelasio-Var", [I]: "Gelasio-Italic-Var", [BI]: "Gelasio-Italic-Var" },
  "Times New Roman": { [R]: "LiberationSerif-Regular", [B]: "LiberationSerif-Bold", [I]: "LiberationSerif-Italic", [BI]: "LiberationSerif-BoldItalic" },
  Arial: { [R]: "LiberationSans-Regular", [B]: "LiberationSans-Bold", [I]: "LiberationSans-Italic", [BI]: "LiberationSans-BoldItalic" },
  "Segoe UI": { [R]: "Selawik-Regular", [B]: "Selawik-Bold", [I]: "Selawik-Regular", [BI]: "Selawik-Bold" },
};
const kind = (bold, italic) => (bold ? (italic ? BI : B) : italic ? I : R);
const fontsDir = new URL("fonts/", document.baseURI);

/** The URL of the stand-in for a family's face, or null for a family without one. */
export function faceUrl(family, bold, italic) {
  const key = Object.keys(families).find((f) => f.toLowerCase() === String(family ?? "").trim().toLowerCase());
  return key ? new URL(`${families[key][kind(bold, italic)]}.woff2`, fontsDir).href : null;
}

/** Register every face under its Office name and load them all (resolves when ready). */
export function loadAllFaces() {
  const faces = [];
  for (const [family, files] of Object.entries(families)) {
    for (const [k, file] of Object.entries(files)) {
      const face = new FontFace(family, `url(${new URL(`${file}.woff2`, fontsDir).href})`, {
        weight: k === B || k === BI ? "700" : "400",
        style: k === I || k === BI ? "italic" : "normal",
      });
      document.fonts.add(face);
      faces.push(face.load().catch(() => null));
    }
  }
  return Promise.all(faces);
}

// The same stand-ins as TrueType files, for the PDF writer (it embeds TrueType;
// Gelasio is cut into static weights). Loaded once, when a document is first opened.
const ttf = {
  Calibri: ["Carlito-Regular", "Carlito-Bold", "Carlito-Italic", "Carlito-BoldItalic"],
  "Calibri Light": ["Carlito-Regular", "Carlito-Bold", "Carlito-Italic", "Carlito-BoldItalic"],
  Cambria: ["Caladea-Regular", "Caladea-Bold", "Caladea-Italic", "Caladea-BoldItalic"],
  Georgia: ["Gelasio-Regular", "Gelasio-Bold", "Gelasio-Italic", "Gelasio-BoldItalic"],
  "Times New Roman": ["LiberationSerif-Regular", "LiberationSerif-Bold", "LiberationSerif-Italic", "LiberationSerif-BoldItalic"],
  Arial: ["LiberationSans-Regular", "LiberationSans-Bold", "LiberationSans-Italic", "LiberationSans-BoldItalic"],
  "Segoe UI": ["Selawik-Regular", "Selawik-Bold", "Selawik-Regular", "Selawik-Bold"],
};
let pdfFontsPromise = null;
/** [{ family, bold, italic, data }] for every face, as the Word and spreadsheet PDF writers take them. */
export function pdfFonts() {
  pdfFontsPromise ??= (async () => {
    const files = new Map();
    const bytes = (name) => {
      if (!files.has(name)) files.set(name, fetch(new URL(`ttf/${name}.ttf`, fontsDir)).then((r) => r.arrayBuffer()).then((b) => new Uint8Array(b)));
      return files.get(name);
    };
    const faces = [];
    for (const [family, names] of Object.entries(ttf)) {
      for (const [i, name] of names.entries()) faces.push({ family, bold: i === 1 || i === 3, italic: i >= 2, data: await bytes(name) });
    }
    return faces;
  })();
  return pdfFontsPromise;
}
