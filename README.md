# rbt-office demo

A static page that opens `.docx` and `.xlsx` files in [rbt-office](https://github.com/fnrbt/rbt-office)'s
own browser shells: five samples, or a file of your own (drag and drop or pick). Edits stay in the
page; the shells' Save downloads the file. Nothing is uploaded.

- `index.html`, `src/`: the page (about a hundred lines around `WordShell` and `createSpreadsheetShell`).
- `public/samples/`: the sample documents and their thumbnails. The macro samples (`.xlsm`, `.docm`) are made by
  rbt-office's `tools/demo-samples/make.fsx`; the page's macro bar enables and runs their VBA.
- `public/fonts/`: free, metrically compatible stand-ins registered under the Office font names
  (Carlito, Caladea, Gelasio, Liberation, Selawik; licences beside them).
- `docs/`: the built site, which GitHub Pages serves.

Build: `npm install`, then `OFFICE_DIST=<rbt-office checkout>/dist npm run build`. `npm run check` serves the
build, opens every sample and reports page errors (`node tools/macros.mjs docs check` runs the macro samples) (`LIVE=<url> npm run check` checks a deployed copy). Both take `BROWSER=firefox` or `BROWSER=webkit` (default
chromium), `DEVICE="iPhone 15"` (a Playwright device profile) and `PLAYWRIGHT=<path to a playwright/index.mjs>` (one whose browsers are installed).

Licence: MIT (see `LICENSE`). Fonts and photos have their own licences: see `public/CREDITS.html`.
