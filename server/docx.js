// DOCX tools, backed by officeparser's structured AST (pure JavaScript).
//
// officeparser loads its optional heavy dependencies (pdf.js for PDFs, tesseract.js for
// OCR) with dynamic import() only when those features are used. DOCX parsing with
// ocr:false touches neither, so the bundle can omit tesseract entirely.

import fs from "node:fs/promises";
import path from "node:path";
import { OfficeParser } from "officeparser";
import { DocCache } from "./cache.js";
import { resolveAllowedFile } from "./paths.js";

const DOCX_EXT = [".docx"];
const MAX_OUTPUT_CHARS = 100_000;

const cache = new DocCache(3);

// ---------- parsing ----------

async function openDocx(inputPath) {
  const file = resolveAllowedFile(inputPath, DOCX_EXT);
  const doc = await cache.get(file, async () => {
    const buffer = await fs.readFile(file.realPath);
    let ast;
    try {
      ast = await OfficeParser.parseOffice(buffer, {
        fileType: "docx",
        ocr: false,
        // Needed for image placeholders (alt text) in the AST; the bytes are dropped below.
        extractAttachments: true,
        ignoreNotes: false,
        ignoreComments: false,
        ignoreHeadersAndFooters: false,
        includeRawContent: false,
      });
    } catch (err) {
      const code = err?.officeIssue?.code;
      if (code === "PASSWORD_REQUIRED" || code === "PASSWORD_INCORRECT") throw new Error("This document is password-protected and cannot be opened.");
      throw new Error(`Could not read DOCX: ${(err?.message || String(err)).replace(/^\[OfficeParser\]:\s*/, "")}`);
    }
    const imageCount = (ast.attachments || []).filter((a) => a.type === "image" && a.mimeType !== "application/octet-stream").length;
    ast.attachments = []; // keep only the structure in memory
    return buildModel(ast, imageCount);
  });
  return { doc, file };
}

// ---------- AST -> model ----------

const isTable = (n) => n?.type === "table";

/** Plain text of a node (no Markdown). */
function plainText(node) {
  if (!node) return "";
  if (node.type === "image") return node.metadata?.altText ? `[image: ${node.metadata.altText}]` : "[image]";
  if (isTable(node)) return tableToGrid(node).rows.map((r) => r.map((c) => c.text).join(" | ")).join("\n");
  if (Array.isArray(node.children) && node.children.length) {
    const sep = node.children.some((c) => c.type === "paragraph" || isTable(c)) ? "\n" : "";
    return node.children.map(plainText).join(sep);
  }
  return node.text || "";
}

/**
 * Inline Markdown for a paragraph-like node. Collects footnotes/endnotes and comments
 * into `ctx` so they can be listed after the body.
 */
function inlineMarkdown(node, ctx) {
  const kids = Array.isArray(node.children) ? node.children : [];
  if (!kids.length) return escapeInline(node.text || "");
  let out = "";
  for (const k of kids) {
    if (k.type === "image") {
      out += k.metadata?.altText ? `[image: ${k.metadata.altText}]` : "[image]";
      continue;
    }
    if (k.type === "break") {
      out += "\n";
      continue;
    }
    if (k.type !== "text") {
      out += k.children?.length ? inlineMarkdown(k, ctx) : escapeInline(k.text || "");
      continue;
    }
    let t = escapeInline(k.text || "");
    const f = k.formatting || {};
    if (t.trim()) {
      const lead = t.match(/^\s*/)[0];
      const trail = t.match(/\s*$/)[0];
      let core = t.trim();
      if (f.bold && f.italic) core = `***${core}***`;
      else if (f.bold) core = `**${core}**`;
      else if (f.italic) core = `*${core}*`;
      if (f.strikethrough) core = `~~${core}~~`;
      const link = k.metadata?.link;
      if (link && k.metadata?.linkType === "external") core = `[${core}](${link})`;
      t = lead + core + trail;
    }
    out += t;
    for (const note of k.notes || []) {
      const kind = note.metadata?.noteType === "endnote" ? "endnote" : "footnote";
      const label = `${kind === "endnote" ? "e" : ""}${note.metadata?.noteId ?? ctx.notes.length + 1}`;
      out += `[^${label}]`;
      ctx.notes.push({ label, kind, text: plainText(note).trim() });
    }
    for (const c of k.comments || []) {
      ctx.comments.push({
        author: c.metadata?.author,
        date: c.metadata?.date,
        on: (k.text || "").trim(),
        text: plainText(c).trim(),
      });
    }
  }
  // Merge adjacent emphasis markers produced by run splitting ("**a****b**" -> "**ab**").
  return out.replace(/\*\*\*\*/g, "").replace(/(?<!\*)\*\*(?!\*)(\s*)\*\*(?!\*)/g, "$1");
}

function escapeInline(s) {
  return s.replace(/\r\n?/g, "\n");
}

/**
 * Expand a table node into a rectangular grid honouring rowSpan/colSpan.
 * Returns { rows: [[{text, merged}]], merges: [{r1,c1,r2,c2,text}], headerRows, nested }.
 */
function tableToGrid(table) {
  const rowNodes = (table.children || []).filter((r) => r.type === "row");
  const grid = [];
  const merges = [];
  let headerRows = 0;
  let nested = 0;
  rowNodes.forEach((row, r) => {
    grid[r] = grid[r] || [];
    let c = 0;
    const cells = (row.children || []).filter((x) => x.type === "cell");
    if (cells.length && cells.every((x) => x.metadata?.style === "header") && headerRows === r) headerRows++;
    for (const cell of cells) {
      while (grid[r][c]) c++;
      const rs = Math.max(1, cell.metadata?.rowSpan || 1);
      const cs = Math.max(1, cell.metadata?.colSpan || 1);
      const parts = [];
      for (const ch of cell.children || []) {
        if (isTable(ch)) {
          nested++;
          parts.push(`[nested table: ${plainText(ch).replace(/\n/g, " / ")}]`);
        } else {
          const t = plainText(ch).trim();
          if (t) parts.push(t);
        }
      }
      const text = parts.length ? parts.join("\n") : (cell.text || "").trim();
      for (let dr = 0; dr < rs; dr++) {
        grid[r + dr] = grid[r + dr] || [];
        for (let dc = 0; dc < cs; dc++) grid[r + dr][c + dc] = dr === 0 && dc === 0 ? { text, merged: false } : { text: "", merged: true };
      }
      if (rs > 1 || cs > 1) merges.push({ r1: r, c1: c, r2: r + rs - 1, c2: c + cs - 1, text });
      c += cs;
    }
  });
  const width = Math.max(0, ...grid.map((r) => r.length));
  const rows = grid.map((r) => Array.from({ length: width }, (_, i) => r[i] || { text: "", merged: false }));
  return { rows, merges, headerRows, nested };
}

/**
 * Walk the AST once and produce ordered blocks with section context.
 * Block: { kind: 'heading'|'paragraph'|'list_item'|'table'|'other', node, text, level?, section, ... }
 */
function buildModel(ast, imageCount) {
  const blocks = [];
  const outline = [];
  let section = [];
  let tableIndex = 0;
  for (const node of ast.content || []) {
    const text = plainText(node).trim();
    if (node.type === "heading") {
      const level = node.metadata?.level || 1;
      section = section.slice(0, level - 1);
      section[level - 1] = text;
      outline.push({ level, text, style: node.metadata?.style });
      blocks.push({ kind: "heading", node, text, level, section: section.filter(Boolean).join(" › ") });
    } else if (node.type === "list") {
      blocks.push({
        kind: "list_item",
        node,
        text,
        depth: node.metadata?.indentation || 0,
        ordered: node.metadata?.listType === "ordered",
        listId: node.metadata?.listId,
        section: section.filter(Boolean).join(" › "),
      });
    } else if (isTable(node)) {
      tableIndex++;
      blocks.push({ kind: "table", node, index: tableIndex, grid: tableToGrid(node), section: section.filter(Boolean).join(" › ") });
    } else {
      blocks.push({ kind: node.type === "paragraph" ? "paragraph" : "other", node, text, section: section.filter(Boolean).join(" › ") });
    }
  }
  const aux = ast.auxiliary || {};
  return {
    ast,
    blocks,
    outline,
    imageCount,
    headers: (aux.headers || []).map(plainText).map((s) => s.trim()).filter(Boolean),
    footers: (aux.footers || []).map(plainText).map((s) => s.trim()).filter(Boolean),
    warnings: (ast.warnings || []).map((w) => w.message).filter(Boolean),
  };
}

// ---------- rendering ----------

function mdCell(s) {
  return s.replace(/\|/g, "\\|").replace(/\t/g, " ").replace(/\n/g, "<br>");
}

function a1(r, c) {
  let s = "";
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return `${s}${r + 1}`;
}

function tableMarkdown(grid) {
  const { rows, merges } = grid;
  if (!rows.length || !rows[0].length) return "(empty table)";
  const line = (r) => `| ${r.map((c) => mdCell(c.text)).join(" | ")} |`;
  const out = [line(rows[0]), `| ${rows[0].map(() => "---").join(" | ")} |`, ...rows.slice(1).map(line)];
  if (merges.length) {
    out.push("", `Merged cells: ${merges.map((m) => `${a1(m.r1, m.c1)}:${a1(m.r2, m.c2)}${m.text ? ` (${JSON.stringify(m.text.length > 40 ? m.text.slice(0, 40) + "…" : m.text)})` : ""}`).join(", ")}`);
  }
  return out.join("\n");
}

function renderDocument(doc) {
  const ctx = { notes: [], comments: [] };
  const parts = [];
  const counters = new Map(); // listId -> [count per depth]
  let prevKind = null;
  for (const b of doc.blocks) {
    if (b.kind === "heading") {
      parts.push(`${"#".repeat(Math.min(6, Math.max(1, b.level)))} ${inlineMarkdown(b.node, ctx).trim()}`);
    } else if (b.kind === "list_item") {
      const key = b.listId ?? "_";
      const levels = counters.get(key) || [];
      levels[b.depth] = (levels[b.depth] || 0) + 1;
      levels.length = b.depth + 1; // reset deeper levels
      counters.set(key, levels);
      const marker = b.ordered ? `${levels[b.depth]}.` : "-";
      const item = `${"    ".repeat(b.depth)}${marker} ${inlineMarkdown(b.node, ctx).trim()}`;
      // Keep list items of one list together (single newline between them).
      if (prevKind === "list_item") parts[parts.length - 1] += `\n${item}`;
      else parts.push(item);
    } else if (b.kind === "table") {
      parts.push(`**Table ${b.index}** (${b.grid.rows.length} rows × ${b.grid.rows[0]?.length || 0} columns)\n\n${tableMarkdown(b.grid)}`);
    } else {
      const t = inlineMarkdown(b.node, ctx).trim();
      if (!t) continue; // skip empty paragraphs
      parts.push(t);
    }
    prevKind = b.kind;
  }
  let body = parts.join("\n\n");
  if (ctx.notes.length) body += `\n\n---\n${ctx.notes.map((n) => `[^${n.label}]: ${n.text}`).join("\n")}`;
  if (ctx.comments.length) {
    body += `\n\n---\nComments:\n${ctx.comments
      .map((c) => `- ${c.author || "Unknown"}${c.date ? ` (${String(c.date).slice(0, 10)})` : ""} on ${JSON.stringify(c.on.length > 60 ? c.on.slice(0, 60) + "…" : c.on)}: ${c.text}`)
      .join("\n")}`;
  }
  if (doc.headers.length || doc.footers.length) {
    body += "\n\n---";
    if (doc.headers.length) body += `\nPage header: ${[...new Set(doc.headers)].join(" | ")}`;
    if (doc.footers.length) body += `\nPage footer: ${[...new Set(doc.footers)].join(" | ")}`;
  }
  return body;
}

// ---------- tools ----------

export async function docxInfo(inputPath) {
  const { doc, file } = await openDocx(inputPath);
  const m = doc.ast.metadata || {};
  const native = m.nativeProperties || {};
  const allText = doc.blocks.map((b) => (b.kind === "table" ? b.grid.rows.flat().map((c) => c.text).join(" ") : b.text)).join("\n");
  const words = (allText.match(/[\p{L}\p{N}][\p{L}\p{N}'’_-]*/gu) || []).length;
  const ctx = { notes: [], comments: [] };
  for (const b of doc.blocks) if (b.kind !== "table") inlineMarkdown(b.node, ctx);
  const tables = doc.blocks.filter((b) => b.kind === "table");
  const pick = (...keys) => keys.map((k) => native[k]).find((v) => v !== undefined && v !== "");

  const result = {
    path: file.realPath,
    file_size_bytes: file.size,
    title: m.title || undefined,
    subject: m.subject || undefined,
    author: m.author || undefined,
    last_modified_by: m.lastModifiedBy || undefined,
    description: m.description || undefined,
    keywords: m.keywords || undefined,
    created: m.created ? new Date(m.created).toISOString() : undefined,
    modified: m.modified ? new Date(m.modified).toISOString() : undefined,
    revision: pick("cp:revision"),
    application: pick("Application", "ep:Application", "app:Application"),
    pages_as_last_saved: pick("Pages", "ep:Pages", "app:Pages"),
    counts: {
      words,
      characters: allText.replace(/\s/g, "").length,
      paragraphs: doc.blocks.filter((b) => b.kind === "paragraph" && b.text).length,
      headings: doc.outline.length,
      list_items: doc.blocks.filter((b) => b.kind === "list_item").length,
      tables: tables.length,
      images: doc.imageCount,
      footnotes: ctx.notes.filter((n) => n.kind === "footnote").length,
      endnotes: ctx.notes.filter((n) => n.kind === "endnote").length,
      comments: ctx.comments.length,
    },
    tables: tables.length
      ? tables.map((t) => ({
          table: t.index,
          rows: t.grid.rows.length,
          columns: t.grid.rows[0]?.length || 0,
          header_rows: t.grid.headerRows || undefined,
          merged_regions: t.grid.merges.length || undefined,
          nested_tables: t.grid.nested || undefined,
          section: t.section || undefined,
        }))
      : undefined,
    outline: doc.outline.length ? doc.outline.slice(0, 200).map((h) => `${"  ".repeat(h.level - 1)}H${h.level}${h.style === "Title" ? " (Title)" : ""}: ${h.text}`) : undefined,
    has_page_headers: doc.headers.length > 0 || undefined,
    has_page_footers: doc.footers.length > 0 || undefined,
    parser_warnings: doc.warnings.length ? [...new Set(doc.warnings)].slice(0, 10) : undefined,
  };
  return JSON.stringify(result, null, 2);
}

export async function docxReadText(inputPath, offset = 0) {
  const { doc, file } = await openDocx(inputPath);
  const full = renderDocument(doc);
  const start = Math.max(0, Math.min(offset, full.length));
  let end = Math.min(full.length, start + MAX_OUTPUT_CHARS);
  if (end < full.length) {
    // Prefer to cut at a paragraph boundary.
    const cut = full.lastIndexOf("\n\n", end);
    if (cut > start + MAX_OUTPUT_CHARS / 2) end = cut;
  }
  const header = `${path.basename(file.realPath)} — ${full.length} characters as Markdown${start > 0 || end < full.length ? `; showing ${start}–${end}` : ""}\n\n`;
  const footer = end < full.length ? `\n\n[Truncated. Call docx_read_text again with offset=${end} to continue.]` : "";
  return header + full.slice(start, end) + footer;
}

export async function docxReadTables(inputPath, table) {
  const { doc, file } = await openDocx(inputPath);
  const tables = doc.blocks.filter((b) => b.kind === "table");
  if (!tables.length) return `${path.basename(file.realPath)} contains no tables.`;
  let selected = tables;
  if (table !== undefined && table !== null) {
    if (!Number.isInteger(table) || table < 1 || table > tables.length) throw new Error(`Table ${table} does not exist (document has ${tables.length} tables).`);
    selected = [tables[table - 1]];
  }
  const parts = [];
  let chars = 0;
  let truncatedAt = null;
  for (const t of selected) {
    const g = t.grid;
    const chunk =
      `## Table ${t.index} of ${tables.length} — ${g.rows.length} rows × ${g.rows[0]?.length || 0} columns` +
      `${t.section ? ` — section: ${t.section}` : ""}${g.headerRows ? ` — ${g.headerRows} header row(s)` : " — no header row marked (first row shown as header)"}\n\n` +
      tableMarkdown(g);
    if (chars + chunk.length > MAX_OUTPUT_CHARS && parts.length) {
      truncatedAt = t.index;
      break;
    }
    parts.push(chunk);
    chars += chunk.length;
  }
  let out = `${path.basename(file.realPath)} — ${tables.length} table(s). Cells are addressed A1-style (column letter, row number). Merged cells appear once, in their top-left cell.\n\n${parts.join("\n\n")}`;
  if (truncatedAt) out += `\n\n[Output truncated. Call docx_read_tables with table=${truncatedAt} (and following) to read the rest.]`;
  return out;
}

export async function docxSearch(inputPath, query, options = {}) {
  const { matchCase = false, maxResults = 50 } = options;
  if (typeof query !== "string" || query === "") throw new Error("Query must be a non-empty string.");
  const { doc, file } = await openDocx(inputPath);
  const q = matchCase ? query : query.toLowerCase();
  const find = (s) => (matchCase ? s : s.toLowerCase()).indexOf(q);
  const snippet = (s, i) => {
    const a = Math.max(0, i - 80);
    const b = Math.min(s.length, i + query.length + 80);
    return `${a > 0 ? "…" : ""}${s.slice(a, b)}${b < s.length ? "…" : ""}`.replace(/\s+/g, " ");
  };

  const hits = [];
  let total = 0;
  const add = (where, s) => {
    const i = find(s);
    if (i < 0) return;
    total++;
    if (hits.length < maxResults) hits.push(`[${where}] ${snippet(s, i)}`);
  };
  const ctx = { notes: [], comments: [] };
  let para = 0;
  let item = 0;
  for (const b of doc.blocks) {
    const sec = b.section ? ` · § ${b.section}` : "";
    if (b.kind === "table") {
      b.grid.rows.forEach((row, r) =>
        row.forEach((cell, c) => {
          if (!cell.merged && cell.text) add(`table ${b.index}, cell ${a1(r, c)} (row ${r + 1}, column ${c + 1})${sec}`, cell.text);
        })
      );
      continue;
    }
    inlineMarkdown(b.node, ctx); // collects notes/comments
    if (b.kind === "heading") add(`heading H${b.level}${sec}`, b.text);
    else if (b.kind === "list_item") add(`list item ${++item}${sec}`, b.text);
    else if (b.text) add(`paragraph ${++para}${sec}`, b.text);
  }
  ctx.notes.forEach((n) => add(`${n.kind} ${n.label}`, n.text));
  ctx.comments.forEach((c) => add(`comment by ${c.author || "Unknown"}`, c.text));
  doc.headers.forEach((h) => add("page header", h));
  doc.footers.forEach((f) => add("page footer", f));

  const head = `${path.basename(file.realPath)} — ${total} match${total === 1 ? "" : "es"} for ${JSON.stringify(query)}${total > hits.length ? ` (showing first ${hits.length})` : ""}`;
  return hits.length ? `${head}\n\n${hits.join("\n")}` : head;
}
