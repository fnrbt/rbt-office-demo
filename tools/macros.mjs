// node tools/macros.mjs [siteDir] [outDir] — serve the built site, open each macro sample, enable its
// macros, run them (from the file's own button where it has one), screenshot, report errors.
// (LIVE=<url> checks a deployed copy.)
import fs from "node:fs"; import path from "node:path"; import http from "node:http";
import { pathToFileURL } from "node:url";
// BROWSER=chromium | firefox | webkit (default chromium); DEVICE=<a Playwright device name, e.g. "iPhone 15">;
// PLAYWRIGHT=<path to another playwright's index.mjs> (one whose browsers are installed).
const playwright = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : "playwright");
const browserName = process.env.BROWSER ?? "chromium";
const browserType = playwright[browserName] ?? playwright.default?.[browserName];
if (!browserType?.launch) throw new Error(`Unknown Playwright browser: ${browserName}`);
const device = process.env.DEVICE ? (playwright.devices ?? playwright.default.devices)[process.env.DEVICE] : null;
if (process.env.DEVICE && !device) throw new Error(`Unknown Playwright device: ${process.env.DEVICE}`);
const site = path.resolve(process.argv[2] ?? "docs"), out = path.resolve(process.argv[3] ?? "check");
fs.mkdirSync(out, { recursive: true });
const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".woff2": "font/woff2", ".png": "image/png" };
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
const b = await browserType.launch();
console.log(`browser: ${browserName} ${b.version()}${device ? ` as ${process.env.DEVICE}` : ""}`);
const errors = [];
const page = await b.newPage(device ? { ...device } : { viewport: { width: 1440, height: 900 } });
page.on("pageerror", (e) => errors.push(String(e).slice(0, 300)));
page.on("console", (m) => { if (m.type() === "error") errors.push("console: " + m.text().slice(0, 300)); });
page.on("dialog", (d) => { errors.push("dialog: " + d.message()); d.dismiss(); });
const idle = () => page.waitForFunction(() => !window.demo.sheet?.grid.MacroRunning, null, { timeout: 120000 });
const enable = async (file) => {
  await page.goto(base + "#" + encodeURIComponent(file));
  // (The samples open with their macros enabled.)
  await page.waitForSelector("#macro-run:not([hidden])", { timeout: 60000 });
};
const shot = async (file) => { await page.screenshot({ path: path.join(out, file + ".png") }); console.log("ran", file); };

await enable("snake.xlsm");
await page.waitForSelector('[data-macro="StartGame"]', { timeout: 20000 });
await page.click('[data-macro="StartGame"]');
for (const key of ["ArrowDown", "ArrowLeft", "ArrowUp", "ArrowRight"]) { await page.waitForTimeout(2500); await page.keyboard.press(key); }
await page.waitForTimeout(1500);
await shot("snake.xlsm");
await page.click('[data-macro="TogglePause"]');

await enable("mandelbrot.xlsm");
await page.waitForSelector('[data-macro="DrawMandelbrot"]', { timeout: 20000 });
await page.click('[data-macro="DrawMandelbrot"]');
await page.waitForTimeout(500); await idle(); await page.waitForTimeout(800);
await shot("mandelbrot.xlsm");

await enable("sierpinski.xlsm");
await page.waitForSelector('[data-macro="main"]', { timeout: 20000 });
await page.click('[data-macro="main"]');
await page.waitForTimeout(500); await idle(); await page.waitForTimeout(1500);
await shot("sierpinski.xlsm");

await enable("rainbow.docm");
await page.click("#macro-run");
await page.waitForTimeout(3000);
await shot("rainbow.docm");

console.log(errors.length ? "ERRORS:\n" + [...new Set(errors)].join("\n") : "no errors");
await b.close(); if (!live) server.close();
