#!/usr/bin/env node
import "./stdio-guard.js";

// Startup is deliberately minimal: only the MCP SDK, zod and the path checks load
// before stdio is served. The PDF/XLSX engines are imported lazily, so a problem
// loading them can never kill the process during Claude's `server/discover` probe
// or the `initialize` handshake — it surfaces as a tool error instead.

import path from "node:path";
import { McpServer } from "@modelcontextprotocol/server";
import { serveStdio } from "@modelcontextprotocol/server/stdio";
import { z } from "zod";
import { setAllowedDirectories } from "./paths.js";
import { MAX_BATCH, MAX_LIST_ENTRIES, createFolder, listDirectory, moveEntries, renameEntry } from "./files.js";

const VERSION = "2.1.0";
const log = (...args) => console.error("[local-documents]", ...args);

process.on("uncaughtException", (err) => log("Uncaught exception:", err));
process.on("unhandledRejection", (err) => log("Unhandled rejection:", err));

// Allowed directories come from the command line (Claude Desktop expands the
// multi-select user_config into separate args) and optionally from an env var.
let roots = [];
try {
  const envDirs = (process.env.LOCAL_DOCUMENTS_ALLOWED_DIRS || "").split(path.delimiter);
  roots = setAllowedDirectories([...process.argv.slice(2), ...envDirs]);
} catch (err) {
  log("Failed to read allowed directories:", err);
}
const runtime = process.versions.bun ? `bun ${process.versions.bun}` : `node ${process.version}`;
log(`v${VERSION} on ${process.platform}/${process.arch}, ${runtime} (${process.execPath})`);
log(`Allowed directories: ${roots.length ? roots.join(", ") : "(none configured)"}`);

// ---------- lazy engine loading ----------

function lazy(name, loader) {
  let promise = null;
  return () => {
    if (!promise) {
      promise = loader().catch((err) => {
        promise = null; // allow a retry on the next call
        log(`Failed to load ${name} engine:`, err);
        throw new Error(`The ${name} engine failed to load on this system (${runtime}, ${process.platform}/${process.arch}): ${err?.message || err}`);
      });
    }
    return promise;
  };
}
const loadPdf = lazy("PDF", () => import("./pdf.js"));
const loadXlsx = lazy("XLSX", () => import("./xlsx.js"));
const loadDocx = lazy("DOCX", () => import("./docx.js"));
const loadXlsxEdit = lazy("XLSX editing", () => import("./xlsx-edit.js"));

// ---------- server ----------

const instructions =
  "Access to local PDF, XLSX and DOCX files inside the user's allowed directories" +
  (roots.length ? ` (${roots.join("; ")})` : "") +
  ". Paths may be absolute or relative to an allowed directory. For PDFs: call pdf_info first, then pdf_read_text; " +
  "if a page has no text layer (scanned), use pdf_render_page to view it. For spreadsheets: call xlsx_info to see sheets " +
  "and used ranges, then xlsx_read_range or xlsx_search. For Word documents: docx_info for the outline, then docx_read_text, " +
  "docx_read_tables or docx_search. To change a workbook, batch all changes into one xlsx_edit call; it writes a NEW workbook by " +
  "default and never changes the original unless overwrite=true is passed; only do that when the user explicitly asks to modify the original. " +
  "To organise files and folders (any type), look first with list_directory (recursive=true for a whole tree), then use create_folder, " +
  "rename, move, or move_batch for many moves at once; these never overwrite " +
  "or delete anything, and a failed move_batch is undone completely.";

const readOnly = { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false };
const organiseAnnotations = { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false };
const editAnnotations = { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: false };
const outputArgs = {
  output_path: z
    .string()
    .optional()
    .describe('Where to save the edited workbook (same extension as the source, .xlsx or .xlsm). Relative paths are relative to the source file\'s folder. Default: a new file next to the source named "<name> (edited).<ext>".'),
  overwrite: z
    .boolean()
    .optional()
    .describe("Allow replacing an existing file. With no output_path, overwrite=true saves over the SOURCE workbook. Default false."),
};
const pathArg = z.string().describe("Absolute path to the file, or a path relative to an allowed directory.");

function wrap(fn) {
  return async (args) => {
    try {
      const out = await fn(args);
      return typeof out === "string" ? { content: [{ type: "text", text: out }] } : out;
    } catch (err) {
      return { isError: true, content: [{ type: "text", text: `Error: ${err?.message || String(err)}` }] };
    }
  };
}

function createServer() {
  const server = new McpServer({ name: "local-documents", version: VERSION }, { capabilities: { tools: {} }, instructions });

  server.registerTool(
    "pdf_info",
    {
      title: "PDF info",
      description: "Get page count, metadata, page sizes, and whether a PDF has a text layer (or is likely scanned).",
      inputSchema: z.object({ path: pathArg }),
      annotations: readOnly,
    },
    wrap(async ({ path: p }) => (await loadPdf()).pdfInfo(p))
  );

  server.registerTool(
    "pdf_read_text",
    {
      title: "Read PDF text",
      description:
        "Extract text from a PDF's existing text layer (no OCR). Output is split by page. Large outputs are truncated with instructions for continuing.",
      inputSchema: z.object({
        path: pathArg,
        pages: z
          .union([z.string(), z.number().int()])
          .optional()
          .describe('Pages to read, 1-based: e.g. "5", "1-10", "2,4,7-9", "20-" (to end). Omit for all pages.'),
      }),
      annotations: readOnly,
    },
    wrap(async ({ path: p, pages }) => (await loadPdf()).pdfReadText(p, pages))
  );

  server.registerTool(
    "pdf_render_page",
    {
      title: "Render PDF page",
      description:
        "Render one PDF page to an image so it can be viewed. Use for scanned pages, charts, diagrams, tables whose layout matters, or pages without a text layer.",
      inputSchema: z.object({
        path: pathArg,
        page: z.coerce.number().int().min(1).describe("1-based page number."),
        dpi: z.coerce.number().min(36).max(400).optional().describe("Resolution (default 144). Longest side is capped at 2200 px."),
        format: z.enum(["auto", "png", "jpeg"]).optional().describe('Image format (default "auto": PNG, or JPEG if the PNG would be very large).'),
      }),
      annotations: readOnly,
    },
    wrap(async ({ path: p, page, dpi, format }) => {
      const { image, caption } = await (await loadPdf()).pdfRenderPage(p, page, dpi ?? 144, format ?? "auto");
      return { content: [{ type: "image", data: image.data, mimeType: image.mimeType }, { type: "text", text: caption }] };
    })
  );

  server.registerTool(
    "xlsx_info",
    {
      title: "XLSX info",
      description: "List worksheets (name, visibility, used range, row/column counts), defined names, and workbook properties.",
      inputSchema: z.object({ path: pathArg }),
      annotations: readOnly,
    },
    wrap(async ({ path: p }) => (await loadXlsx()).xlsxInfo(p))
  );

  server.registerTool(
    "xlsx_read_range",
    {
      title: "Read XLSX range",
      description:
        "Read cell values from a worksheet range as a tab-separated grid with row numbers and column letters. Values are the cached results stored in the file.",
      inputSchema: z.object({
        path: pathArg,
        sheet: z.string().describe("Sheet name (case-insensitive) or 1-based sheet index as a string."),
        range: z.string().describe('A1-style range, e.g. "A1:F50", "B7", "A:D" (whole columns), "10:20" (whole rows).'),
        include_formulas: z.boolean().optional().describe("Also list the formulas of formula cells in the range (default false)."),
      }),
      annotations: readOnly,
    },
    wrap(async ({ path: p, sheet, range, include_formulas }) => (await loadXlsx()).xlsxReadRange(p, sheet, range, include_formulas ?? false))
  );

  server.registerTool(
    "xlsx_search",
    {
      title: "Search XLSX",
      description: "Find cells whose value contains the query text (case-insensitive by default) across all sheets or one sheet.",
      inputSchema: z.object({
        path: pathArg,
        query: z.string().min(1).describe("Text to search for."),
        sheet: z.string().optional().describe("Limit the search to one sheet (name or 1-based index)."),
        match_case: z.boolean().optional().describe("Case-sensitive match (default false)."),
        whole_cell: z.boolean().optional().describe("Match the entire cell value only (default false)."),
        include_formulas: z.boolean().optional().describe("Also search formula text (default false)."),
        max_results: z.coerce.number().int().min(1).max(1000).optional().describe("Maximum matches to return (default 100)."),
      }),
      annotations: readOnly,
    },
    wrap(async ({ path: p, query, sheet, match_case, whole_cell, include_formulas, max_results }) =>
      (await loadXlsx()).xlsxSearch(p, query, {
        sheet,
        matchCase: match_case ?? false,
        wholeCell: whole_cell ?? false,
        includeFormulas: include_formulas ?? false,
        maxResults: max_results ?? 100,
      })
    )
  );

  const cellValue = z.union([z.string(), z.number(), z.boolean(), z.null(), z.object({ date: z.string().describe('ISO date "YYYY-MM-DD" or "YYYY-MM-DDTHH:MM[:SS]"') })]);
  const cellEdit = z.strictObject({
    sheet: z.string().describe("Sheet name (case-insensitive) or 1-based index as a string."),
    cell: z.string().describe('Single cell in A1 style, e.g. "G3".'),
    value: cellValue.optional().describe('New value: number, string, boolean, null (clears) or {"date": "YYYY-MM-DD"}. Strings are always text, even if they start with "=".'),
    formula: z.string().optional().describe('Formula in English Excel syntax, with or without the leading "=", e.g. "COUNTA(Sheet1!A:A)".'),
    clear: z.literal(true).optional().describe("Clear the cell's value/formula (its formatting is kept)."),
  });
  const addSheetEdit = z.strictObject({ add_sheet: z.string().describe("Name of a new, empty worksheet to append. Cell edits in the same call may target it.") });

  server.registerTool(
    "xlsx_edit",
    {
      title: "Edit XLSX",
      description:
        "Apply a batch of edits to an .xlsx/.xlsm workbook in ONE transaction: set cell values, set formulas, clear cells, and add worksheets. " +
        "Each cell edit has exactly one of value / formula / clear. The workbook is edited surgically: only the affected XML elements change, " +
        "cell formatting is kept, and every unrelated part (charts, pivots, macros, add-ins, customXml, custom properties, styles) is copied byte-for-byte. " +
        "By default the result is saved as a NEW file next to the source; the original is only replaced with overwrite=true. " +
        "If any edit is invalid, nothing is written. Formulas are stored, not evaluated; Excel recalculates on open.",
      inputSchema: z.object({
        path: pathArg,
        edits: z
          .array(z.union([cellEdit, addSheetEdit]))
          .min(1)
          .describe('Edits applied in order (add_sheet edits first). Example: [{"sheet":"Sheet1","cell":"G3","value":177.5},{"sheet":"Summary","cell":"B2","formula":"COUNTA(Sheet1!A:A)"},{"add_sheet":"Notes"}]'),
        ...outputArgs,
      }),
      annotations: editAnnotations,
    },
    wrap(async ({ path: p, edits, output_path, overwrite }) => (await loadXlsxEdit()).xlsxEdit(p, edits, { outputPath: output_path, overwrite: overwrite ?? false }))
  );

  server.registerTool(
    "docx_info",
    {
      title: "DOCX info",
      description:
        "Get a Word document's properties (title, author, dates), heading outline, table list, and counts of words, paragraphs, lists, tables, images, footnotes and comments.",
      inputSchema: z.object({ path: pathArg }),
      annotations: readOnly,
    },
    wrap(async ({ path: p }) => (await loadDocx()).docxInfo(p))
  );

  server.registerTool(
    "docx_read_text",
    {
      title: "Read DOCX text",
      description:
        "Read a Word document as Markdown in reading order: headings, paragraphs, bulleted/numbered lists (with nesting), tables, footnotes, comments and page headers/footers. Long documents are paginated with an offset.",
      inputSchema: z.object({
        path: pathArg,
        offset: z.coerce.number().int().min(0).optional().describe("Character offset to continue from, as given in a previous truncated response (default 0)."),
      }),
      annotations: readOnly,
    },
    wrap(async ({ path: p, offset }) => (await loadDocx()).docxReadText(p, offset ?? 0))
  );

  server.registerTool(
    "docx_read_tables",
    {
      title: "Read DOCX tables",
      description: "Extract the tables of a Word document as Markdown grids, with merged cells, header rows and the section each table is in.",
      inputSchema: z.object({
        path: pathArg,
        table: z.coerce.number().int().min(1).optional().describe("1-based table number to read a single table. Omit for all tables."),
      }),
      annotations: readOnly,
    },
    wrap(async ({ path: p, table }) => (await loadDocx()).docxReadTables(p, table))
  );

  server.registerTool(
    "docx_search",
    {
      title: "Search DOCX",
      description:
        "Find text in a Word document (paragraphs, headings, list items, table cells, footnotes, comments, headers/footers). Returns each match's location and surrounding text.",
      inputSchema: z.object({
        path: pathArg,
        query: z.string().min(1).describe("Text to search for."),
        match_case: z.boolean().optional().describe("Case-sensitive match (default false)."),
        max_results: z.coerce.number().int().min(1).max(1000).optional().describe("Maximum matches to return (default 50)."),
      }),
      annotations: readOnly,
    },
    wrap(async ({ path: p, query, match_case, max_results }) =>
      (await loadDocx()).docxSearch(p, query, { matchCase: match_case ?? false, maxResults: max_results ?? 50 })
    )
  );

  const entryArg = z.string().describe("Absolute path to the file or folder, or a path relative to an allowed directory.");
  const destArg = z
    .string()
    .describe(
      'New path for the item, including its name. End it with "/" to move the item INTO that folder and keep its name. Relative paths are relative to the allowed directory the source is in.'
    );
  const createFoldersArg = z.boolean().optional().describe("Create missing destination folders (default false).");

  server.registerTool(
    "list_directory",
    {
      title: "List folder",
      description:
        "List the files and folders in a folder, with sizes and modification times. With no path, lists every allowed directory. " +
        "Paths in the output are relative to the allowed directory and can be passed to move, rename and move_batch as they are. " +
        "Hidden files are skipped by default; symbolic links are shown but not followed.",
      inputSchema: z.object({
        path: z.string().optional().describe("Folder to list: absolute, or relative to an allowed directory. Omit to list all allowed directories."),
        recursive: z.boolean().optional().describe("Also list subfolders (default false)."),
        max_depth: z.coerce.number().int().min(1).max(20).optional().describe("With recursive=true, how many levels to list (default 5)."),
        include_hidden: z.boolean().optional().describe('Include hidden files such as ".DS_Store" (default false).'),
        max_entries: z.coerce.number().int().min(1).max(MAX_LIST_ENTRIES).optional().describe("Maximum entries to return (default 1000)."),
      }),
      annotations: readOnly,
    },
    wrap(async ({ path: p, recursive, max_depth, include_hidden, max_entries }) =>
      listDirectory(p, { recursive: recursive ?? false, maxDepth: max_depth ?? 5, includeHidden: include_hidden ?? false, maxEntries: max_entries ?? 1000 })
    )
  );

  server.registerTool(
    "create_folder",
    {
      title: "Create folder",
      description: "Create a folder inside an allowed directory. Succeeds without change if the folder already exists.",
      inputSchema: z.object({
        path: z.string().describe("Absolute path of the new folder, or a path relative to the first allowed directory."),
        parents: z.boolean().optional().describe("Also create missing parent folders (default false)."),
      }),
      annotations: { ...organiseAnnotations, idempotentHint: true },
    },
    wrap(async ({ path: p, parents }) => createFolder(p, { parents: parents ?? false }))
  );

  server.registerTool(
    "rename",
    {
      title: "Rename",
      description: "Rename a file or folder in place (same folder). Fails if an item with the new name already exists; nothing is overwritten.",
      inputSchema: z.object({
        path: entryArg,
        new_name: z.string().min(1).describe('The new name only, e.g. "2024 Budget.xlsx" (no folders).'),
      }),
      annotations: organiseAnnotations,
    },
    wrap(async ({ path: p, new_name }) => renameEntry(p, new_name))
  );

  server.registerTool(
    "move",
    {
      title: "Move",
      description:
        "Move (and optionally rename) one file or folder of any type within the allowed directories. Fails if the destination exists; nothing is overwritten. " +
        "For several items, use move_batch instead.",
      inputSchema: z.object({ from: entryArg, to: destArg, create_folders: createFoldersArg }),
      annotations: organiseAnnotations,
    },
    wrap(async ({ from, to, create_folders }) => moveEntries([{ from, to }], { createFolders: create_folders ?? false }))
  );

  server.registerTool(
    "move_batch",
    {
      title: "Move many",
      description:
        "Move and/or rename many files and folders in ONE all-or-nothing call, e.g. to reorganise a folder. Moves run in order, and each sees the " +
        "result of the ones before it. If any move fails, every earlier move is undone and the folders created for them are removed. " +
        "Nothing is ever overwritten or deleted.",
      inputSchema: z.object({
        moves: z
          .array(z.strictObject({ from: entryArg, to: destArg }))
          .min(1)
          .max(MAX_BATCH)
          .describe('Example: [{"from":"Inbox/invoice-march.pdf","to":"Finance/2024/Invoices/"},{"from":"Inbox/notes.docx","to":"Projects/Alpha/Meeting notes.docx"}]'),
        create_folders: createFoldersArg,
      }),
      annotations: organiseAnnotations,
    },
    wrap(async ({ moves, create_folders }) => moveEntries(moves, { createFolders: create_folders ?? false }))
  );

  return server;
}

// serveStdio owns protocol-era negotiation: it answers the 2026-07-28 `server/discover`
// probe and still serves the 2025-era `initialize` handshake from the same factory.
serveStdio(createServer, { onerror: (err) => log("Transport error:", err?.message || err) });

// Warm the engines in the background once stdio is up, so the first tool call is fast.
setTimeout(() => {
  loadPdf().catch(() => {});
  loadXlsx().catch(() => {});
  loadDocx().catch(() => {});
}, 50).unref();
