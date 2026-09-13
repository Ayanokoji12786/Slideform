import { cp, mkdir, rm } from "node:fs/promises";

await rm("dist", { force: true, recursive: true });
await mkdir("dist/vendor", { recursive: true });
await cp("src", "dist", { recursive: true });
await cp("node_modules/jszip/dist/jszip.min.js", "dist/vendor/jszip.min.js");
await cp("node_modules/exceljs/dist/exceljs.min.js", "dist/vendor/exceljs.min.js");
await cp("node_modules/pdfjs-dist/build/pdf.min.mjs", "dist/vendor/pdf.min.mjs");
await cp("node_modules/pdfjs-dist/build/pdf.worker.min.mjs", "dist/vendor/pdf.worker.min.mjs");
