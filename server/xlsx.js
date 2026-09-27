// XLSX tools, backed by ExcelJS (pure JavaScript).

import path from "node:path";
import ExcelJS from "exceljs";
import { DocCache } from "./cache.js";
import { resolveAllowedFile } from "./paths.js";

const XLSX_EXT = [".xlsx", ".xlsm"];
const MAX_RANGE_CELLS = 20_000;
const MAX_OUTPUT_CHARS = 100_000;
const EXCEL_MAX_ROWS = 1_048_576;
const EXCEL_MAX_COLS = 16_384;

const cache = new DocCache(3);

async function openWorkbook(inputPath) {
  const file = resolveAllowedFile(inputPath, XLSX_EXT);
  const wb = await cache.get(file, async () => {
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.readFile(file.realPath);
    } catch (err) {
      throw new Error(`Could not read workbook: ${err?.message || err}`);
    }
    return workbook;
  });
  return { wb, file };
}

// ---------- A1 helpers ----------

export function colToNum(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Parse an A1-style range. Supports "B2", "A1:D20", "A:C", "3:10", "$A$1:$B$2" and a
 * "Sheet!" prefix (which is ignored — the sheet argument wins).
 * Returns { r1, c1, r2, c2 } (1-based, inclusive). Open-ended rows/cols are clipped
 * to the sheet's used area by the caller via `bounds`.
 */
export function parseRange(range, bounds) {
  let s = String(range).trim();
  const bang = s.lastIndexOf("!");
  if (bang >= 0) s = s.slice(bang + 1);
  s = s.replace(/\$/g, "").toUpperCase();
  const parts = s.split(":");
  if (parts.length > 2 || !parts[0]) throw new Error(`Invalid range "${range}". Use e.g. "A1:D20", "B5", "A:C" or "3:10".`);
  const one = (p) => {
    let m;
    if ((m = p.match(/^([A-Z]{1,3})(\d+)$/))) return { c: colToNum(m[1]), r: parseInt(m[2], 10) };
    if ((m = p.match(/^([A-Z]{1,3})$/))) return { c: colToNum(m[1]), r: null };
    if ((m = p.match(/^(\d+)$/))) return { c: null, r: parseInt(m[1], 10) };
    throw new Error(`Invalid cell reference "${p}" in range "${range}".`);
  };
  const a = one(parts[0]);
  const b = parts.length === 2 ? one(parts[1]) : a;
  if ((a.r === null) !== (b.r === null) || (a.c === null) !== (b.c === null)) {
    throw new Error(`Invalid range "${range}": both ends must be the same kind (cells, columns, or rows).`);
  }
  let r1 = a.r ?? 1;
  let r2 = b.r ?? Math.max(bounds.rows, 1);
  let c1 = a.c ?? 1;
  let c2 = b.c ?? Math.max(bounds.cols, 1);
  if (r1 > r2) [r1, r2] = [r2, r1];
  if (c1 > c2) [c1, c2] = [c2, c1];
  if (r1 < 1 || c1 < 1 || r2 > EXCEL_MAX_ROWS || c2 > EXCEL_MAX_COLS) throw new Error(`Range "${range}" is outside Excel's limits.`);
  return { r1, c1, r2, c2 };
}

const addr = (r, c) => `${numToCol(c)}${r}`;

// ---------- cell value formatting ----------

function formatDate(d) {
  if (Number.isNaN(d.getTime())) return "";
  const iso = d.toISOString();
  // ExcelJS returns dates as UTC; drop the time part for date-only values.
  return iso.endsWith("T00:00:00.000Z") ? iso.slice(0, 10) : iso.replace(".000Z", "Z");
}

function scalarToString(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date) return formatDate(v);
  if (typeof v === "number") return Number.isInteger(v) ? String(v) : String(+v.toPrecision(15));
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "string") return v;
  if (typeof v === "object") {
    if (Array.isArray(v.richText)) return v.richText.map((t) => t.text).join("");
    if (v.error) return String(v.error);
    if ("text" in v && ("hyperlink" in v || typeof v.text === "string")) return scalarToString(v.text);
    if ("result" in v) return scalarToString(v.result);
    if ("formula" in v || "sharedFormula" in v) return "";
  }
  return String(v);
}

/** Returns { text, formula } for a cell. */
function cellInfo(cell) {
  const v = cell.value;
  let formula;
  if (v && typeof v === "object" && !(v instanceof Date)) {
    if (typeof v.formula === "string") formula = v.formula;
    else if (typeof v.sharedFormula === "string") {
      try {
        formula = cell.formula || undefined;
      } catch {}
      formula = formula || `(shared from ${v.sharedFormula})`;
    }
  }
  return { text: scalarToString(v), formula };
}

function isMergedSlave(cell) {
  return cell.isMerged && cell.master && cell.master.address !== cell.address;
}

function escapeCell(s) {
  return s.replace(/\\/g, "\\\\").replace(/\t/g, "\\t").replace(/\r?\n/g, "\\n");
}

// ---------- sheet lookup ----------

export function findSheet(wb, sheet) {
  const sheets = wb.worksheets;
  if (sheet === undefined || sheet === null || sheet === "") {
    if (!sheets.length) throw new Error("Workbook has no worksheets.");
    return sheets[0];
  }
  const s = String(sheet);
  let ws = sheets.find((w) => w.name === s) || sheets.find((w) => w.name.toLowerCase() === s.toLowerCase());
  if (!ws && /^\d+$/.test(s)) ws = sheets[parseInt(s, 10) - 1];
  if (!ws) throw new Error(`Sheet "${s}" not found. Available sheets: ${sheets.map((w) => JSON.stringify(w.name)).join(", ")}`);
  return ws;
}

function usedBounds(ws) {
  return { rows: ws.rowCount || 0, cols: ws.columnCount || 0 };
}

// ---------- tools ----------

export async function xlsxInfo(inputPath) {
  const { wb, file } = await openWorkbook(inputPath);
  const sheets = wb.worksheets.map((ws, i) => {
    const { rows, cols } = usedBounds(ws);
    const merges = ws.model?.merges || [];
    return {
      index: i + 1,
      name: ws.name,
      state: ws.state || "visible",
      used_range: rows && cols ? `A1:${addr(rows, cols)}` : null,
      dimensions: ws.dimensions?.range ?? undefined,
      rows_with_data: ws.actualRowCount,
      max_row: rows,
      max_column: cols,
      merged_ranges: merges.length || undefined,
      tables: ws.tables && Object.keys(ws.tables).length ? Object.keys(ws.tables) : undefined,
    };
  });
  let definedNames;
  try {
    const model = wb.definedNames.model || [];
    definedNames = model.length ? model.map((d) => ({ name: d.name, refers_to: d.ranges })) : undefined;
  } catch {}
  const result = {
    path: file.realPath,
    file_size_bytes: file.size,
    sheet_count: sheets.length,
    sheets,
    defined_names: definedNames,
    creator: wb.creator || undefined,
    last_modified_by: wb.lastModifiedBy || undefined,
    created: wb.created instanceof Date ? wb.created.toISOString() : undefined,
    modified: wb.modified instanceof Date ? wb.modified.toISOString() : undefined,
  };
  return JSON.stringify(result, null, 2);
}

export async function xlsxReadRange(inputPath, sheet, range, includeFormulas = false) {
  const { wb, file } = await openWorkbook(inputPath);
  const ws = findSheet(wb, sheet);
  const bounds = usedBounds(ws);
  if (!range || !String(range).trim()) {
    if (!bounds.rows || !bounds.cols) return `Sheet "${ws.name}" is empty.`;
    range = `A1:${addr(bounds.rows, bounds.cols)}`;
  }
  const req = parseRange(range, bounds);
  // Never read past the used area; there is nothing there.
  const r2 = Math.min(req.r2, Math.max(bounds.rows, req.r1));
  const c2 = Math.min(req.c2, Math.max(bounds.cols, req.c1));
  const { r1, c1 } = req;
  const width = c2 - c1 + 1;
  let lastRow = r2;
  let truncatedReason = null;
  if ((r2 - r1 + 1) * width > MAX_RANGE_CELLS) {
    lastRow = r1 + Math.max(1, Math.floor(MAX_RANGE_CELLS / width)) - 1;
    truncatedReason = `cell limit (${MAX_RANGE_CELLS})`;
  }

  const lines = [];
  const formulas = [];
  lines.push(["", ...Array.from({ length: width }, (_, i) => numToCol(c1 + i))].join("\t"));
  let chars = lines[0].length;
  let r;
  for (r = r1; r <= lastRow; r++) {
    const row = ws.getRow(r);
    const vals = [];
    for (let c = c1; c <= c2; c++) {
      const cell = row.getCell(c);
      if (isMergedSlave(cell)) {
        vals.push("");
        continue;
      }
      const { text, formula } = cellInfo(cell);
      vals.push(escapeCell(text));
      if (includeFormulas && formula) formulas.push(`${addr(r, c)}: =${formula}  →  ${text}`);
    }
    const line = `${r}\t${vals.join("\t")}`;
    chars += line.length + 1;
    if (chars > MAX_OUTPUT_CHARS) {
      truncatedReason = `output size limit (~${MAX_OUTPUT_CHARS} chars)`;
      break;
    }
    lines.push(line);
  }
  const shownLast = truncatedReason ? r - 1 : r2;

  const header =
    `${path.basename(file.realPath)} — sheet "${ws.name}", range ${addr(r1, c1)}:${addr(shownLast, c2)}` +
    (req.r2 > r2 || req.c2 > c2 ? ` (requested ${String(range).trim()}, clipped to used area)` : "") +
    `\nTab-separated. First row = column letters, first column = row numbers. Values are cached results (formulas are not recalculated); \\t and \\n inside cells are escaped.\n\n`;
  let out = header + lines.join("\n");
  if (formulas.length) out += `\n\nFormulas:\n${formulas.join("\n")}`;
  if (truncatedReason) out += `\n\n[Truncated at row ${shownLast} due to ${truncatedReason}. Continue with range "${addr(shownLast + 1, c1)}:${addr(r2, c2)}".]`;
  return out;
}

export async function xlsxSearch(inputPath, query, options = {}) {
  const { sheet, matchCase = false, wholeCell = false, includeFormulas = false, maxResults = 100 } = options;
  if (typeof query !== "string" || query === "") throw new Error("Query must be a non-empty string.");
  const { wb, file } = await openWorkbook(inputPath);
  const sheets = sheet !== undefined && sheet !== null && sheet !== "" ? [findSheet(wb, sheet)] : wb.worksheets;
  const q = matchCase ? query : query.toLowerCase();
  const test = (s) => {
    if (!s) return false;
    const t = matchCase ? s : s.toLowerCase();
    return wholeCell ? t.trim() === q.trim() : t.includes(q);
  };

  const hits = [];
  let total = 0;
  for (const ws of sheets) {
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      row.eachCell({ includeEmpty: false }, (cell, c) => {
        if (isMergedSlave(cell)) return;
        const { text, formula } = cellInfo(cell);
        const inValue = test(text);
        const inFormula = includeFormulas && formula && test(formula);
        if (!inValue && !inFormula) return;
        total++;
        if (hits.length < maxResults) {
          const shown = text.length > 300 ? text.slice(0, 300) + "…" : text;
          hits.push(`${ws.name}!${addr(r, c)}\t${escapeCell(shown)}${inFormula && formula ? `\t(formula: =${formula})` : ""}`);
        }
      });
    });
  }

  const head = `${path.basename(file.realPath)} — ${total} match${total === 1 ? "" : "es"} for ${JSON.stringify(query)}${
    sheets.length === 1 ? ` in sheet "${sheets[0].name}"` : ` across ${sheets.length} sheets`
  }${total > hits.length ? ` (showing first ${hits.length})` : ""}`;
  return hits.length ? `${head}\n\n${hits.join("\n")}` : head;
}
