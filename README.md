# rbt-office demo

A static page that opens `.docx` and `.xlsx` files in [rbt-office](https://github.com/fnrbt/rbt-office)'s
own browser shells: five samples, or a file of your own (drag and drop or pick). Edits stay in the
page; the shells' Save downloads the file. Nothing is uploaded.

- `index.html`, `src/`: the page (about a hundred lines around `WordShell` and `createSpreadsheetShell`).
- `public/samples/`: the sample documents and their thumbnails.
- `public/fonts/`: free, metrically compatible stand-ins registered under the Office font names
  (Carlito, Caladea, Gelasio, Liberation, Selawik; licences beside them).
- `docs/`: the built site, which GitHub Pages serves.

Build: `OFFICE_DIST=<rbt-office checkout>/dist ./build.sh`. `node tools/check.mjs docs out` serves the
build, opens every sample and reports page errors.
