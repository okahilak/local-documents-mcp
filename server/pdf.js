// PDF tools.
//   - Text and metadata: pdf.js (legacy build for Node) reading the existing text layer.
//     pdf.js needs DOMMatrix/ImageData/Path2D, provided by @napi-rs/canvas.
//   - Rendering: PDFium compiled to WebAssembly, encoded to PNG/JPEG with pure-JS encoders.

// Must stay the first import: installs the globals pdf.js evaluates at load time.
import "./pdfjs-polyfills.js";
import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";
import { PDFiumLibrary } from "@hyzyla/pdfium";
import { PNG } from "pngjs";
import jpeg from "jpeg-js";
import { DocCache } from "./cache.js";
import { resolveAllowedFile } from "./paths.js";

const require = createRequire(import.meta.url);
const PDFJS_ROOT = path.dirname(require.resolve("pdfjs-dist/package.json"));
/**
 * pdf.js validates cMapUrl/standardFontDataUrl (and wasmUrl/iccUrl) by requiring a trailing
 * "/", then in Node reads files with fs.readFile(dirUrl + name). A Windows path ends in "\",
 * which fails that check, so use forward slashes (fs accepts them on Windows too).
 */
export function toPdfjsDirUrl(dir) {
  return dir.replace(/\\/g, "/").replace(/\/*$/, "/");
}
const CMAP_URL = toPdfjsDirUrl(path.join(PDFJS_ROOT, "cmaps"));
const STANDARD_FONTS_URL = toPdfjsDirUrl(path.join(PDFJS_ROOT, "standard_fonts"));
pdfjs.GlobalWorkerOptions.workerSrc = pathToFileURL(path.join(PDFJS_ROOT, "legacy", "build", "pdf.worker.mjs")).href;

const PDF_EXT = [".pdf"];
const MAX_TEXT_CHARS = 100_000;
const MAX_RENDER_PIXELS = 2200; // longest side, keeps images readable but reasonably small
const PNG_SIZE_LIMIT = 3.5 * 1024 * 1024;

const pdfjsCache = new DocCache(4, (doc) => doc.destroy().catch(() => {}));
const pdfiumCache = new DocCache(2, (doc) => {
  try {
    doc.destroy();
  } catch {}
});

let pdfiumLibPromise = null;
function getPdfium() {
  if (!pdfiumLibPromise) {
    const wasmPath = require.resolve("@hyzyla/pdfium/pdfium.wasm");
    pdfiumLibPromise = fs
      .readFile(wasmPath)
      .then((buf) => PDFiumLibrary.init({ wasmBinary: buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) }));
    pdfiumLibPromise.catch(() => (pdfiumLibPromise = null));
  }
  return pdfiumLibPromise;
}

async function openPdfjs(file) {
  return pdfjsCache.get(file, async () => {
    const data = new Uint8Array(await fs.readFile(file.realPath));
    const task = pdfjs.getDocument({
      data,
      cMapUrl: CMAP_URL,
      cMapPacked: true,
      standardFontDataUrl: STANDARD_FONTS_URL,
      isEvalSupported: false,
      disableFontFace: true,
      useSystemFonts: false,
      stopAtErrors: false,
      verbosity: 0,
    });
    try {
      return await task.promise;
    } catch (err) {
      if (err?.name === "PasswordException") throw new Error("This PDF is password-protected and cannot be opened.");
      throw new Error(`Could not parse PDF: ${err?.message || err}`);
    }
  });
}

async function openPdfium(file) {
  return pdfiumCache.get(file, async () => {
    const lib = await getPdfium();
    const data = new Uint8Array(await fs.readFile(file.realPath));
    try {
      return await lib.loadDocument(data);
    } catch (err) {
      throw new Error(`Could not open PDF for rendering: ${err?.message || err}`);
    }
  });
}

/** Parse "1-3,5,8-" style page specs into a sorted list of 1-based page numbers. */
export function parsePages(spec, numPages) {
  if (spec === undefined || spec === null || spec === "" || spec === "all") {
    return Array.from({ length: numPages }, (_, i) => i + 1);
  }
  if (typeof spec === "number") spec = String(spec);
  if (Array.isArray(spec)) spec = spec.join(",");
  const out = new Set();
  for (const partRaw of String(spec).split(",")) {
    const part = partRaw.trim();
    if (!part) continue;
    const m = part.match(/^(\d*)\s*-\s*(\d*)$/);
    let a, b;
    if (m) {
      a = m[1] ? parseInt(m[1], 10) : 1;
      b = m[2] ? parseInt(m[2], 10) : numPages;
    } else if (/^\d+$/.test(part)) {
      a = b = parseInt(part, 10);
    } else {
      throw new Error(`Invalid page spec "${part}". Use e.g. "3", "1-5", "2,4,6-8", or "10-".`);
    }
    if (a < 1 || b < a) throw new Error(`Invalid page range "${part}".`);
    if (a > numPages) throw new Error(`Page ${a} is out of range (document has ${numPages} pages).`);
    for (let i = a; i <= Math.min(b, numPages); i++) out.add(i);
  }
  if (out.size === 0) throw new Error("No pages selected.");
  return [...out].sort((x, y) => x - y);
}

/** Rebuild readable lines from pdf.js text items. */
function textItemsToString(items) {
  let out = "";
  let lastY = null;
  let lastEndX = null;
  for (const item of items) {
    if (typeof item.str !== "string") continue;
    if (item.str === "") {
      // Empty items are end-of-line markers; their position belongs to the next line.
      if (item.hasEOL && out && !out.endsWith("\n")) out += "\n";
      lastEndX = null;
      continue;
    }
    const [, , , , x, y] = item.transform;
    const fontSize = Math.hypot(item.transform[2], item.transform[3]) || 10;
    if (lastY !== null && Math.abs(y - lastY) > fontSize * 0.5) {
      if (!out.endsWith("\n")) out += "\n";
      // A big vertical gap usually means a new paragraph.
      if (Math.abs(y - lastY) > fontSize * 1.8) out += "\n";
      lastEndX = null;
    } else if (lastEndX !== null && x - lastEndX > fontSize * 0.25 && !out.endsWith(" ") && !out.endsWith("\n") && item.str && !item.str.startsWith(" ")) {
      out += " ";
    }
    out += item.str;
    if (item.hasEOL) {
      out += "\n";
      lastEndX = null;
    } else {
      lastEndX = x + (item.width || 0);
    }
    lastY = y;
  }
  return out.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}

async function pageText(doc, n) {
  const page = await doc.getPage(n);
  try {
    const content = await page.getTextContent({ includeMarkedContent: false, disableNormalization: false });
    return textItemsToString(content.items);
  } finally {
    page.cleanup();
  }
}

function fmtDate(s) {
  if (!s) return undefined;
  const d = pdfjs.PDFDateString.toDateObject(s);
  return d ? d.toISOString() : s;
}

export async function pdfInfo(inputPath) {
  const file = resolveAllowedFile(inputPath, PDF_EXT);
  const doc = await openPdfjs(file);
  const meta = await doc.getMetadata().catch(() => ({ info: {} }));
  const info = meta.info || {};

  const numPages = doc.numPages;
  const sizes = [];
  const sampleCount = Math.min(numPages, 50);
  for (let i = 1; i <= sampleCount; i++) {
    const page = await doc.getPage(i);
    const vp = page.getViewport({ scale: 1 });
    sizes.push({ page: i, width_pt: Math.round(vp.width), height_pt: Math.round(vp.height), rotation: page.rotate });
    page.cleanup();
  }
  // Distinct sizes only, to keep output compact.
  const distinctSizes = [];
  for (const s of sizes) {
    const last = distinctSizes[distinctSizes.length - 1];
    if (last && last.width_pt === s.width_pt && last.height_pt === s.height_pt && last.rotation === s.rotation) last.to_page = s.page;
    else distinctSizes.push({ from_page: s.page, to_page: s.page, width_pt: s.width_pt, height_pt: s.height_pt, rotation: s.rotation });
  }

  // Probe the text layer on up to 5 pages spread through the document.
  const probe = [...new Set([1, Math.ceil(numPages / 4), Math.ceil(numPages / 2), Math.ceil((3 * numPages) / 4), numPages])].filter((n) => n >= 1);
  const textChars = {};
  for (const n of probe) textChars[n] = (await pageText(doc, n)).replace(/\s/g, "").length;
  const pagesWithText = Object.values(textChars).filter((c) => c > 20).length;

  let outlineCount = 0;
  try {
    const outline = await doc.getOutline();
    outlineCount = outline ? outline.length : 0;
  } catch {}

  const result = {
    path: file.realPath,
    file_size_bytes: file.size,
    pages: numPages,
    pdf_version: info.PDFFormatVersion,
    title: info.Title || undefined,
    author: info.Author || undefined,
    subject: info.Subject || undefined,
    keywords: info.Keywords || undefined,
    creator: info.Creator || undefined,
    producer: info.Producer || undefined,
    created: fmtDate(info.CreationDate),
    modified: fmtDate(info.ModDate),
    is_encrypted: Boolean(info.EncryptFilterName) || undefined,
    has_acroform: info.IsAcroFormPresent || undefined,
    top_level_bookmarks: outlineCount || undefined,
    page_sizes: sampleCount < numPages ? { first_50_pages: distinctSizes } : distinctSizes,
    text_layer_probe: { non_whitespace_chars_by_page: textChars },
    text_layer:
      pagesWithText === probe.length ? "present" : pagesWithText === 0 ? "absent (likely scanned — use pdf_render_page)" : "partial (some pages may be scanned — use pdf_render_page for those)",
  };
  return JSON.stringify(result, null, 2);
}

export async function pdfReadText(inputPath, pages) {
  const file = resolveAllowedFile(inputPath, PDF_EXT);
  const doc = await openPdfjs(file);
  const pageList = parsePages(pages, doc.numPages);

  const parts = [];
  let total = 0;
  let truncatedAt = null;
  const emptyPages = [];
  for (const n of pageList) {
    const text = await pageText(doc, n);
    if (text.replace(/\s/g, "").length === 0) emptyPages.push(n);
    const chunk = `--- Page ${n} ---\n${text || "[no text layer on this page — use pdf_render_page to view it]"}\n`;
    if (total + chunk.length > MAX_TEXT_CHARS && parts.length > 0) {
      truncatedAt = n;
      break;
    }
    parts.push(chunk);
    total += chunk.length;
  }

  let header = `${path.basename(file.realPath)} — ${doc.numPages} pages total; showing ${truncatedAt ? `${pageList[0]}–${truncatedAt - 1}` : pages ? `pages ${pages}` : "all pages"}\n\n`;
  let footer = "";
  if (truncatedAt) {
    const remaining = pageList.filter((p) => p >= truncatedAt);
    footer = `\n[Output truncated at ~${MAX_TEXT_CHARS} characters. Remaining requested pages start at ${truncatedAt} (${remaining.length} pages). Call pdf_read_text again with pages="${truncatedAt}-${remaining[remaining.length - 1]}".]`;
  }
  if (emptyPages.length) footer += `\n[Pages without a text layer: ${emptyPages.join(", ")}. They may be scanned images; use pdf_render_page to view them.]`;
  return header + parts.join("\n") + footer;
}

export async function pdfRenderPage(inputPath, page, dpi = 144, format = "auto") {
  const file = resolveAllowedFile(inputPath, PDF_EXT);
  const doc = await openPdfium(file);
  const numPages = doc.getPageCount();
  if (!Number.isInteger(page) || page < 1 || page > numPages) {
    throw new Error(`Page ${page} is out of range (document has ${numPages} pages).`);
  }
  const p = doc.getPage(page - 1);
  const { originalWidth, originalHeight } = p.getOriginalSize();
  let scale = Math.max(0.1, Math.min(dpi, 400) / 72);
  const longest = Math.max(originalWidth, originalHeight) * scale;
  if (longest > MAX_RENDER_PIXELS) scale *= MAX_RENDER_PIXELS / longest;

  const rendered = await p.render({
    scale,
    colorSpace: "BGRA", // with REVERSE_BYTE_ORDER (set by the library) this is RGBA in memory
    renderFormFields: true,
    transparent: false,
    render: async ({ data }) => data,
  });
  const { width, height, data } = rendered;

  let mimeType;
  let bytes;
  if (format === "jpeg") {
    ({ mimeType, bytes } = encodeJpeg(data, width, height));
  } else {
    const png = new PNG({ width, height, colorType: 6, inputHasAlpha: true });
    png.data = Buffer.from(data.buffer, data.byteOffset, data.byteLength);
    bytes = PNG.sync.write(png, { colorType: 2, deflateLevel: 6 });
    mimeType = "image/png";
    if (format === "auto" && bytes.length > PNG_SIZE_LIMIT) ({ mimeType, bytes } = encodeJpeg(data, width, height));
  }

  return {
    image: { data: Buffer.from(bytes).toString("base64"), mimeType },
    caption: `${path.basename(file.realPath)} — page ${page} of ${numPages} (${width}×${height}px, ${mimeType}, ${Math.round(bytes.length / 1024)} KB)`,
  };
}

function encodeJpeg(data, width, height) {
  const out = jpeg.encode({ data: Buffer.from(data.buffer, data.byteOffset, data.byteLength), width, height }, 85);
  return { mimeType: "image/jpeg", bytes: out.data };
}
