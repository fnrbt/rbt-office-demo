// node tools/check.mjs [siteDir] [outDir] — serve the built site, open each sample, screenshot, report errors.
import fs from "node:fs"; import path from "node:path"; import http from "node:http";
import { chromium } from "playwright";
const site = path.resolve(process.argv[2] ?? "site"), out = path.resolve(process.argv[3] ?? "check");
fs.mkdirSync(out, { recursive: true });
const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".woff2": "font/woff2", ".png": "image/png", ".dat": "application/octet-stream" };
const server = http.createServer((req, res) => {
  let p = path.join(site, decodeURIComponent(new URL(req.url, "http://x").pathname));
  if (fs.existsSync(p) && fs.statSync(p).isDirectory()) p = path.join(p, "index.html");
  if (!fs.existsSync(p)) { res.statusCode = 404; return res.end(); }
  res.setHeader("Content-Type", types[path.extname(p)] ?? "application/octet-stream");
  fs.createReadStream(p).pipe(res);
});
const live = process.env.LIVE;
if (!live) await new Promise((r) => server.listen(0, "127.0.0.1", r));
const base = live ?? `http://127.0.0.1:${server.address().port}/`;
const b = await chromium.launch();
const errors = [];
const page = await b.newPage({ viewport: { width: 1440, height: 900 } });
page.on("pageerror", (e) => errors.push(String(e).slice(0, 300)));
page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text().slice(0, 300)); });
page.on("requestfailed", (r) => errors.push("failed: " + r.url()));
page.on("response", (r) => { if (r.status() >= 400) errors.push(`${r.status()} ${r.url()}`); });
await page.goto(base);
await page.waitForTimeout(1500);
await page.screenshot({ path: path.join(out, "home.png") });
for (const f of process.argv.slice(4).length ? process.argv.slice(4) : ["annual-review.docx", "field-notes.docx", "services-agreement.docx", "dashboard.xlsx", "budget.xlsx"]) {
  await page.goto(base + "#" + encodeURIComponent(f));
  await page.waitForTimeout(f.endsWith(".docx") ? 4500 : 3500);
  await page.screenshot({ path: path.join(out, f + ".png") });
  console.log("opened", f, await page.evaluate(() => document.querySelector("#appname")?.textContent));
}
console.log(errors.length ? "ERRORS:\n" + [...new Set(errors)].join("\n") : "no errors");
await b.close(); if (!live) server.close();
