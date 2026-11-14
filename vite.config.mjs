// Build: OFFICE_DIST=<rbt-office checkout>/dist node <vite> build  -> site/
import path from "node:path";
const office = process.env.OFFICE_DIST;
if (!office) throw new Error("set OFFICE_DIST to an rbt-office dist folder");
export default {
  base: "./",
  resolve: { alias: [{ find: /^@rbt\/office\/(.*)$/, replacement: path.join(office, "$1") }] },
  build: { outDir: "docs", emptyOutDir: true, chunkSizeWarningLimit: 8000, target: "es2022" },
};
