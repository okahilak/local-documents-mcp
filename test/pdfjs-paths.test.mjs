// Regression test for Windows: pdf.js requires cMapUrl/standardFontDataUrl to end in "/",
// and a Windows path built with path.join ends in "\". Runs on any OS by using path.win32.
// SERVER_ENTRY env var can point at an unpacked bundle's server/index.js.
import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const serverDir = path.dirname(path.resolve(process.env.SERVER_ENTRY || path.join(here, "..", "server", "index.js")));
const { toPdfjsDirUrl } = await import(pathToFileURL(path.join(serverDir, "pdf.js")).href);
// Same pdf.js the server uses (bundle or dev tree).
const pdfjs = await import(pathToFileURL(path.join(serverDir, "..", "node_modules", "pdfjs-dist", "legacy", "build", "pdf.mjs")).href);

const winRoot = "C:\\Users\\Jane Doe\\AppData\\Roaming\\Claude\\Claude Extensions\\local-documents\\node_modules\\pdfjs-dist";
const oldStyle = path.win32.join(winRoot, "cmaps") + path.win32.sep; // what v1.2.1 produced on Windows

test("toPdfjsDirUrl converts Windows paths to forward slashes with a trailing slash", () => {
  assert.equal(
    toPdfjsDirUrl(path.win32.join(winRoot, "cmaps")),
    "C:/Users/Jane Doe/AppData/Roaming/Claude/Claude Extensions/local-documents/node_modules/pdfjs-dist/cmaps/"
  );
  assert.equal(toPdfjsDirUrl(oldStyle), toPdfjsDirUrl(path.win32.join(winRoot, "cmaps")));
  assert.equal(toPdfjsDirUrl("\\\\server\\share\\ext\\pdfjs-dist\\standard_fonts"), "//server/share/ext/pdfjs-dist/standard_fonts/");
  assert.equal(toPdfjsDirUrl("/Users/me/ext/node_modules/pdfjs-dist/cmaps"), "/Users/me/ext/node_modules/pdfjs-dist/cmaps/");
  assert.equal(toPdfjsDirUrl("/already/has/slash/"), "/already/has/slash/");
});

test("pdf.js rejects the old backslash form and accepts the fixed form", async () => {
  const data = () => new Uint8Array(fs.readFileSync(path.join(here, "fixtures", "allowed", "report.pdf")));
  const opts = (dir) => ({ data: data(), verbosity: 0, cMapUrl: dir, cMapPacked: true, standardFontDataUrl: dir });
  assert.throws(() => pdfjs.getDocument(opts(oldStyle)), /Invalid factory url: .* must include trailing slash/);
  const task = pdfjs.getDocument(opts(toPdfjsDirUrl(path.win32.join(winRoot, "cmaps"))));
  await task.destroy();
});

test("the server's real cMap/standard-font directories resolve to readable files", () => {
  const root = path.join(serverDir, "..", "node_modules", "pdfjs-dist");
  const cmaps = toPdfjsDirUrl(path.join(root, "cmaps"));
  const fonts = toPdfjsDirUrl(path.join(root, "standard_fonts"));
  // pdf.js's Node reader does fs.readFile(dirUrl + name).
  assert.ok(fs.readFileSync(cmaps + "UniGB-UCS2-H.bcmap").length > 0);
  assert.ok(fs.readdirSync(fonts).length > 0);
});
