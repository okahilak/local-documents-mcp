// Transactional, surgical XLSX/XLSM editing.
//
// The workbook is treated as a ZIP/OOXML package:
//   - Only the XML parts an edit needs are modified (the edited worksheets, and when
//     required workbook.xml, sharedStrings.xml, calcChain.xml, workbook.xml.rels and
//     [Content_Types].xml). Inside those parts only the affected elements are rewritten.
//   - Every other ZIP entry is copied as raw bytes (see ooxml/zip.js).
//   - ExcelJS is never used to save; it is only used to read the result back as a check.
//
// Safety: all edits are validated before anything is written; the result is written to a
// temp file, re-read and validated, then atomically renamed into place. Writes are
// serialized per file (in-process queue + lock file), and replacing an existing file is a
// compare-and-swap against the bytes that were read.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { readZip, readEntry, writeZip } from "./ooxml/zip.js";
import * as X from "./ooxml/xml.js";
import { normalizeFormula, translateFormula, colToNum, numToCol } from "./ooxml/formula.js";
import { resolveAllowedFile, resolveAllowedOutput, assertOutputStillAllowed, samePath, PathError } from "./paths.js";

const EDIT_EXT = [".xlsx", ".xlsm"];
const MAX_EDITS = 10_000;
const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL = {
  officeDocument: NS_REL + "/officeDocument",
  worksheet: NS_REL + "/worksheet",
  sharedStrings: NS_REL + "/sharedStrings",
  calcChain: NS_REL + "/calcChain",
  table: NS_REL + "/table",
};
const CT_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const WORKBOOK_AFTER_CALCPR = ["oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"];

// ---------- small helpers ----------

const sha256 = (buf) => crypto.createHash("sha256").update(buf).digest("hex");

function parseCellRef(ref, what = "cell") {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(ref ?? "").trim());
  if (!m) throw new Error(`Invalid ${what} reference ${JSON.stringify(ref)}; use A1 style like "G3".`);
  const col = colToNum(m[1]);
  const row = parseInt(m[2], 10);
  if (col < 1 || col > 16384 || row < 1 || row > 1048576) throw new Error(`${what} ${ref} is outside Excel's sheet limits.`);
  return { row, col, ref: `${numToCol(col)}${row}` };
}

function parseRangeRef(ref) {
  const [a, b] = String(ref).split(":");
  const p = parseCellRef(a, "range");
  const q = b ? parseCellRef(b, "range") : p;
  return { r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) };
}

const inRange = (g, row, col) => row >= g.r1 && row <= g.r2 && col >= g.c1 && col <= g.c2;

function partDir(partName) {
  return partName.includes("/") ? partName.slice(0, partName.lastIndexOf("/")) : "";
}

function relsPathFor(partName) {
  const dir = partDir(partName);
  const base = partName.slice(dir ? dir.length + 1 : 0);
  return `${dir ? dir + "/" : ""}_rels/${base}.rels`;
}

function resolveTarget(fromPart, target) {
  let t = target;
  try {
    t = decodeURI(target);
  } catch {}
  if (t.startsWith("/")) return path.posix.normalize(t.slice(1));
  return path.posix.normalize(path.posix.join(partDir(fromPart), t));
}

// ---------- package access ----------

class Package {
  constructor(buf) {
    this.zip = readZip(buf);
    this.byName = new Map(this.zip.entries.map((e) => [e.name.toLowerCase(), e]));
    this.texts = new Map(); // lower name -> current text (only parts we read)
    this.changed = new Set(); // lower names modified
    this.deleted = new Set();
    this.additions = []; // { name, text }
  }
  has(name) {
    const k = name.toLowerCase();
    return (this.byName.has(k) && !this.deleted.has(k)) || this.additions.some((a) => a.name.toLowerCase() === k);
  }
  text(name) {
    const k = name.toLowerCase();
    if (this.texts.has(k)) return this.texts.get(k);
    const e = this.byName.get(k);
    if (!e) throw new Error(`Package part not found: ${name}`);
    const raw = readEntry(e);
    if ((raw[0] === 0xff && raw[1] === 0xfe) || (raw[0] === 0xfe && raw[1] === 0xff)) throw new Error(`Part ${name} is UTF-16 encoded, which is not supported for editing.`);
    const s = raw.toString("utf8");
    this.texts.set(k, s);
    return s;
  }
  set(name, text) {
    const k = name.toLowerCase();
    const added = this.additions.find((a) => a.name.toLowerCase() === k);
    if (added) added.text = text;
    else this.changed.add(k);
    this.texts.set(k, text);
  }
  add(name, text) {
    if (this.has(name)) throw new Error(`Part already exists: ${name}`);
    this.additions.push({ name, text });
    this.texts.set(name.toLowerCase(), text);
  }
  remove(name) {
    this.deleted.add(name.toLowerCase());
  }
  build() {
    const changes = new Map();
    for (const k of this.changed) changes.set(k, Buffer.from(this.texts.get(k), "utf8"));
    for (const k of this.deleted) changes.set(k, null);
    return writeZip(this.zip, changes, this.additions.map((a) => ({ name: a.name, data: Buffer.from(a.text, "utf8") })));
  }
  modifiedPartNames() {
    return [...this.changed].filter((k) => !this.deleted.has(k)).map((k) => this.byName.get(k).name);
  }
}

function readRels(pkg, partName) {
  const rp = relsPathFor(partName);
  if (!pkg.has(rp)) return { path: rp, rels: [] };
  const xml = pkg.text(rp);
  const root = X.rootElement(xml);
  const rels = X.scanChildren(xml, root.openEnd, root.closeStart)
    .filter((e) => e.local === "Relationship")
    .map((e) => {
      const a = X.parseAttrs(xml, e);
      return { el: e, id: X.decodeEntities(X.getAttr(a, "Id") || ""), type: X.getAttr(a, "Type"), target: X.decodeEntities(X.getAttr(a, "Target") || ""), external: X.getAttr(a, "TargetMode") === "External" };
    });
  return { path: rp, rels };
}

// ---------- workbook model ----------

function loadWorkbook(pkg) {
  const pkgRels = readRels(pkg, "").rels; // package-level _rels/.rels
  const office = pkgRels.find((r) => r.type === REL.officeDocument || r.type?.endsWith("/officeDocument"));
  if (!office) throw new Error("Not a spreadsheet package: no officeDocument relationship.");
  if (office.type.startsWith("http://purl.oclc.org/ooxml/")) throw new Error("Strict Open XML workbooks are not supported for editing.");
  const wbPath = resolveTarget("", office.target);
  const xml = pkg.text(wbPath);
  const wbRoot = X.rootElement(xml);
  if (wbRoot.local !== "workbook") throw new Error("The main document part is not a workbook.");
  if (X.prefixFor(xml, wbRoot, NS_MAIN) === null) throw new Error("Unsupported workbook namespace (only Transitional SpreadsheetML is supported).");
  const kids = X.scanChildren(xml, wbRoot.openEnd, wbRoot.closeStart);
  const sheetsEl = kids.find((k) => k.local === "sheets");
  if (!sheetsEl) throw new Error("Workbook has no <sheets> element.");
  const wbRels = readRels(pkg, wbPath);
  const relById = new Map(wbRels.rels.map((r) => [r.id, r]));
  const sheets = X.scanChildren(xml, sheetsEl.openEnd, sheetsEl.closeStart)
    .filter((e) => e.local === "sheet")
    .map((e, index) => {
      const a = X.parseAttrs(xml, e);
      const rid = a.find((x) => x.name === "r:id" || /:id$/.test(x.name))?.value;
      const rel = relById.get(X.decodeEntities(rid || ""));
      return {
        index,
        name: X.decodeEntities(X.getAttr(a, "name") || ""),
        sheetId: X.getAttr(a, "sheetId"),
        rel,
        isWorksheet: rel?.type === REL.worksheet,
        path: rel && !rel.external ? resolveTarget(wbPath, rel.target) : null,
      };
    });
  const workbookPr = kids.find((k) => k.local === "workbookPr");
  const date1904 = workbookPr ? /^(1|true)$/.test(X.getAttr(X.parseAttrs(xml, workbookPr), "date1904") || "") : false;
  const sstRel = wbRels.rels.find((r) => r.type === REL.sharedStrings);
  const calcRel = wbRels.rels.find((r) => r.type === REL.calcChain);
  return {
    path: wbPath,
    rels: wbRels,
    sheets,
    date1904,
    sstPath: sstRel ? resolveTarget(wbPath, sstRel.target) : null,
    calcChainPath: calcRel ? resolveTarget(wbPath, calcRel.target) : null,
    calcChainRel: calcRel,
  };
}

function findSheet(wb, name) {
  const s = String(name ?? "");
  let sh = wb.sheets.find((x) => x.name === s) || wb.sheets.find((x) => x.name.toLowerCase() === s.toLowerCase());
  if (!sh && /^\d+$/.test(s)) sh = wb.sheets[parseInt(s, 10) - 1];
  if (!sh) throw new Error(`Sheet "${s}" not found. Available sheets: ${wb.sheets.map((x) => JSON.stringify(x.name)).join(", ")}`);
  if (!sh.isWorksheet || !sh.path) throw new Error(`Sheet "${sh.name}" is not a worksheet (chart sheets and dialog sheets cannot be edited).`);
  return sh;
}

export function validateSheetName(name, existing) {
  if (typeof name !== "string") throw new Error("Sheet name must be a string.");
  if (name.length < 1 || name.length > 31) throw new Error("Sheet names must be 1–31 characters long.");
  if (/[:\\/?*[\]]/.test(name)) throw new Error("Sheet names cannot contain : \\ / ? * [ or ].");
  if (/[\x00-\x1F]/.test(name)) throw new Error("Sheet names cannot contain control characters.");
  if (name.startsWith("'") || name.endsWith("'")) throw new Error("Sheet names cannot start or end with an apostrophe.");
  if (name.trim().toLowerCase() === "history") throw new Error('"History" is a reserved sheet name in Excel.');
  if (existing.some((n) => n.toLowerCase() === name.toLowerCase())) throw new Error(`A sheet named "${name}" already exists (sheet names are case-insensitive).`);
}

// ---------- add sheet ----------

function addSheet(pkg, wb, name) {
  validateSheetName(name, wb.sheets.map((s) => s.name));
  // Part name
  const firstWs = wb.sheets.find((s) => s.isWorksheet && s.path);
  const dir = firstWs ? partDir(firstWs.path) : `${partDir(wb.path) ? partDir(wb.path) + "/" : ""}worksheets`;
  let n = wb.sheets.length + 1;
  while (pkg.has(`${dir}/sheet${n}.xml`)) n++;
  const partName = `${dir}/sheet${n}.xml`;

  // Relationship
  const relsXml = pkg.has(wb.rels.path) ? pkg.text(wb.rels.path) : null;
  const usedIds = new Set(wb.rels.rels.map((r) => r.id));
  let k = wb.rels.rels.length + 1;
  while (usedIds.has(`rId${k}`)) k++;
  const rId = `rId${k}`;
  const target = path.posix.relative(partDir(wb.path), partName);
  if (relsXml) {
    const root = X.rootElement(relsXml);
    const el = X.withPrefix(root.name, "Relationship");
    const ins = `<${el} Id="${rId}" Type="${REL.worksheet}" Target="${X.escapeAttr(target)}"/>`;
    const at = root.selfClosing ? null : root.closeStart;
    pkg.set(wb.rels.path, at === null ? relsXml.slice(0, root.start) + `<${root.name}${relsXml.slice(root.start + root.name.length + 1, root.openEnd - 2)}>${ins}</${root.name}>` + relsXml.slice(root.end) : relsXml.slice(0, at) + ins + relsXml.slice(at));
  } else {
    pkg.add(wb.rels.path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="${rId}" Type="${REL.worksheet}" Target="${X.escapeAttr(target)}"/></Relationships>`);
  }
  wb.rels.rels.push({ id: rId, type: REL.worksheet, target });

  // Content type override
  const ctXml = pkg.text("[Content_Types].xml");
  const ctRoot = X.rootElement(ctXml);
  const ov = `<${X.withPrefix(ctRoot.name, "Override")} PartName="/${X.escapeAttr(partName)}" ContentType="${CT_WORKSHEET}"/>`;
  pkg.set("[Content_Types].xml", ctXml.slice(0, ctRoot.closeStart) + ov + ctXml.slice(ctRoot.closeStart));

  // <sheet> entry in workbook.xml
  const xml = pkg.text(wb.path);
  const wbRoot = X.rootElement(xml);
  const sheetsEl = X.scanChildren(xml, wbRoot.openEnd, wbRoot.closeStart).find((e) => e.local === "sheets");
  const ids = wb.sheets.map((s) => parseInt(s.sheetId, 10)).filter(Number.isFinite);
  const sheetId = (ids.length ? Math.max(...ids) : 0) + 1;
  let rPrefix = X.prefixFor(xml, wbRoot, NS_REL);
  let nsDecl = "";
  if (rPrefix === null || rPrefix === "") {
    rPrefix = "r";
    nsDecl = ` xmlns:r="${NS_REL}"`;
  }
  const sheetEl = `<${X.withPrefix(sheetsEl.name, "sheet")} name="${X.escapeAttr(name)}" sheetId="${sheetId}" ${rPrefix}:id="${rId}"${nsDecl}/>`;
  pkg.set(wb.path, xml.slice(0, sheetsEl.closeStart) + sheetEl + xml.slice(sheetsEl.closeStart));

  // Minimal worksheet part
  pkg.add(
    partName,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><dimension ref="A1"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`
  );
  const sheet = { index: wb.sheets.length, name, sheetId: String(sheetId), rel: { id: rId, type: REL.worksheet }, isWorksheet: true, path: partName, added: true };
  wb.sheets.push(sheet);
  return sheet;
}

// ---------- shared strings ----------

class SharedStrings {
  constructor(pkg, sstPath) {
    this.pkg = pkg;
    this.path = sstPath;
    this.loaded = false;
    this.pending = [];
    this.refs = 0;
  }
  load() {
    if (this.loaded) return;
    const xml = this.pkg.text(this.path);
    const root = X.rootElement(xml);
    const items = root.selfClosing ? [] : X.scanChildren(xml, root.openEnd, root.closeStart).filter((e) => e.local === "si");
    this.count = items.length;
    this.index = new Map();
    items.forEach((si, i) => {
      if (si.selfClosing) return;
      const kids = X.scanChildren(xml, si.openEnd, si.closeStart);
      if (kids.length === 1 && kids[0].local === "t") {
        const t = kids[0];
        const inner = t.selfClosing ? "" : xml.slice(t.openEnd, t.closeStart);
        if (!inner.includes("<")) {
          const key = X.decodeEntities(inner);
          if (!this.index.has(key)) this.index.set(key, i);
        }
      }
    });
    this.loaded = true;
  }
  indexOf(str) {
    this.load();
    const key = X.encodeXstring(str);
    this.refs++;
    if (this.index.has(key)) return { index: this.index.get(key), added: false };
    const index = this.count + this.pending.length;
    this.pending.push(key);
    this.index.set(key, index);
    return { index, added: true };
  }
  /** Append new <si> items and bump counts — only if something was added. */
  finalize() {
    if (!this.pending.length) return false;
    const xml = this.pkg.text(this.path);
    const root = X.rootElement(xml);
    const si = X.withPrefix(root.name, "si");
    const t = X.withPrefix(root.name, "t");
    const items = this.pending
      .map((s) => `<${si}><${t}${/^\s|\s$/.test(s) ? ' xml:space="preserve"' : ""}>${X.escapeText(s)}</${t}></${si}>`)
      .join("");
    const attrs = X.parseAttrs(xml, root);
    const bump = (name, by) => {
      const v = X.getAttr(attrs, name);
      if (v !== undefined && /^\d+$/.test(v)) X.setAttr(attrs, name, String(parseInt(v, 10) + by));
    };
    bump("count", this.refs);
    bump("uniqueCount", this.pending.length);
    const open = X.startTag(root.name, attrs);
    const body = root.selfClosing ? "" : xml.slice(root.openEnd, root.closeStart);
    this.pkg.set(this.path, xml.slice(0, root.start) + open + body + items + `</${root.name}>` + xml.slice(root.end));
    return true;
  }
}

// ---------- worksheet editing ----------

function excelSerial(dateStr, date1904) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?$/.exec(dateStr);
  if (!m) throw new Error(`Invalid date ${JSON.stringify(dateStr)}; use "YYYY-MM-DD" or "YYYY-MM-DDTHH:MM[:SS]".`);
  const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), 0) + Math.round(parseFloat(m[6] || "0") * 1000);
  const d = new Date(ms);
  if (d.getUTCFullYear() !== +m[1] || d.getUTCMonth() !== +m[2] - 1 || d.getUTCDate() !== +m[3]) throw new Error(`Invalid date ${JSON.stringify(dateStr)}.`);
  let serial = (ms - Date.UTC(1899, 11, 30)) / 86400000;
  if (date1904) serial -= 1462;
  if (serial < (date1904 ? 0 : 1)) throw new Error(`Date ${dateStr} is before Excel's first date.`);
  return serial;
}

function numberText(n) {
  if (!Number.isFinite(n)) throw new Error("Numbers must be finite.");
  if (Object.is(n, -0)) return "0";
  return String(n);
}

/** Collect <f> info for all formula cells in sheetData (only called when needed). */
function scanFormulaCells(xml, rows) {
  const out = [];
  for (const row of rows) {
    if (row.selfClosing) continue;
    for (const c of X.scanChildren(xml, row.openEnd, row.closeStart)) {
      if (c.local !== "c" || c.selfClosing) continue;
      const f = X.scanChildren(xml, c.openEnd, c.closeStart).find((k) => k.local === "f");
      if (!f) continue;
      const cref = X.getAttr(X.parseAttrs(xml, c), "r");
      if (!cref) continue;
      const fa = X.parseAttrs(xml, f);
      out.push({ cell: parseCellRef(cref), cellEl: c, f, attrs: fa, text: f.selfClosing ? "" : xml.slice(f.openEnd, f.closeStart) });
    }
  }
  return out;
}

function sheetTables(pkg, sheet) {
  const { rels } = readRels(pkg, sheet.path);
  const out = [];
  for (const r of rels.filter((x) => x.type === REL.table && !x.external)) {
    const p = resolveTarget(sheet.path, r.target);
    if (!pkg.has(p)) continue;
    const xml = pkg.text(p);
    const root = X.rootElement(xml);
    const a = X.parseAttrs(xml, root);
    const ref = X.getAttr(a, "ref");
    if (!ref) continue;
    const hdr = X.getAttr(a, "headerRowCount");
    out.push({ name: X.decodeEntities(X.getAttr(a, "displayName") || X.getAttr(a, "name") || p), range: parseRangeRef(ref), headerRows: hdr === undefined ? 1 : parseInt(hdr, 10) });
  }
  return out;
}

/**
 * Apply cell edits to one worksheet part. Returns notes and cells whose formulas were
 * removed (for calcChain cleanup).
 */
function editWorksheet(pkg, wb, sheet, cellEdits, sst) {
  let xml = pkg.text(sheet.path);
  const root = X.rootElement(xml);
  if (root.local !== "worksheet") throw new Error(`Part ${sheet.path} is not a worksheet.`);
  if (X.prefixFor(xml, root, NS_MAIN) === null) throw new Error(`Sheet "${sheet.name}" uses an unsupported namespace (Strict Open XML?).`);
  const kids = X.scanChildren(xml, root.openEnd, root.closeStart);
  const sd = kids.find((k) => k.local === "sheetData");
  if (!sd) throw new Error(`Sheet "${sheet.name}" has no <sheetData>.`);
  const notes = [];

  // --- Guard rails: merged cells, array/data-table formulas, table headers.
  const merges = [];
  const mc = kids.find((k) => k.local === "mergeCells");
  if (mc && !mc.selfClosing) {
    for (const m of X.scanChildren(xml, mc.openEnd, mc.closeStart)) {
      const ref = X.getAttr(X.parseAttrs(xml, m), "ref");
      if (ref) merges.push({ ref, ...parseRangeRef(ref) });
    }
  }
  const tables = sheetTables(pkg, sheet);
  let rows = sd.selfClosing ? [] : X.scanChildren(xml, sd.openEnd, sd.closeStart).filter((e) => e.local === "row");

  // Rows without an explicit r are positioned implicitly; make them explicit first so
  // inserting rows cannot shift them.
  if (rows.some((r) => X.getAttr(X.parseAttrs(xml, r), "r") === undefined)) {
    let prev = 0;
    let out = xml.slice(0, sd.openEnd);
    let pos = sd.openEnd;
    for (const r of rows) {
      const a = X.parseAttrs(xml, r);
      const rr = X.getAttr(a, "r");
      const num = rr === undefined ? prev + 1 : parseInt(rr, 10);
      prev = num;
      out += xml.slice(pos, r.start);
      out += rr === undefined ? X.startTag(r.name, [{ name: "r", value: String(num), quote: '"' }, ...a], r.selfClosing) : xml.slice(r.start, r.openEnd);
      pos = r.openEnd;
    }
    xml = out + xml.slice(pos);
    return editWorksheetAgain();
  }
  function editWorksheetAgain() {
    pkg.set(sheet.path, xml);
    return editWorksheet(pkg, wb, sheet, cellEdits, sst);
  }

  const formulaCells = scanFormulaCells(xml, rows);
  const blockers = formulaCells
    .filter((fc) => ["array", "dataTable"].includes(X.getAttr(fc.attrs, "t")) && X.getAttr(fc.attrs, "ref"))
    .map((fc) => ({ kind: X.getAttr(fc.attrs, "t") === "array" ? "an array formula" : "a data table", ref: X.getAttr(fc.attrs, "ref"), ...parseRangeRef(X.getAttr(fc.attrs, "ref")) }));
  const formulaAt = new Map(formulaCells.map((fc) => [fc.cell.ref, fc]));

  for (const e of cellEdits) {
    const { row, col, ref } = e.pos;
    const m = merges.find((g) => inRange(g, row, col) && !(g.r1 === row && g.c1 === col));
    if (m) throw new Error(`${sheet.name}!${ref} is inside merged range ${m.ref}; edit its top-left cell ${numToCol(m.c1)}${m.r1} instead.`);
    const b = blockers.find((g) => inRange(g, row, col));
    if (b) throw new Error(`${sheet.name}!${ref} is part of ${b.kind} (${b.ref}); editing it is not supported.`);
    const t = tables.find((tb) => tb.headerRows > 0 && row >= tb.range.r1 && row < tb.range.r1 + tb.headerRows && col >= tb.range.c1 && col <= tb.range.c2);
    if (t) throw new Error(`${sheet.name}!${ref} is a header cell of table "${t.name}"; renaming table columns is not supported.`);
  }

  // --- Shared formulas whose master is being replaced: promote the first remaining
  // dependent to master (with translated formula and a fresh ref).
  const targetRefs = new Set(cellEdits.map((e) => e.pos.ref));
  const fRewrites = new Map(); // cellRef -> new <f> element text
  for (const e of cellEdits) {
    const fc = formulaAt.get(e.pos.ref);
    if (!fc || X.getAttr(fc.attrs, "t") !== "shared" || !X.getAttr(fc.attrs, "ref") || !fc.text) continue;
    const si = X.getAttr(fc.attrs, "si");
    const deps = formulaCells
      .filter((d) => d !== fc && X.getAttr(d.attrs, "t") === "shared" && X.getAttr(d.attrs, "si") === si && !d.text && !targetRefs.has(d.cell.ref))
      .sort((a, b) => a.cell.row - b.cell.row || a.cell.col - b.cell.col);
    if (!deps.length) continue;
    const nm = deps[0];
    const bbox = deps.reduce((g, d) => ({ r1: Math.min(g.r1, d.cell.row), c1: Math.min(g.c1, d.cell.col), r2: Math.max(g.r2, d.cell.row), c2: Math.max(g.c2, d.cell.col) }), { r1: Infinity, c1: Infinity, r2: 0, c2: 0 });
    const formula = translateFormula(X.decodeEntities(fc.text), nm.cell.row - fc.cell.row, nm.cell.col - fc.cell.col);
    let attrs = X.setAttr([...nm.attrs], "ref", `${numToCol(bbox.c1)}${bbox.r1}:${numToCol(bbox.c2)}${bbox.r2}`);
    fRewrites.set(nm.cell.ref, `${X.startTag(nm.f.name, attrs)}${X.escapeText(formula)}</${nm.f.name}>`);
    notes.push(`${sheet.name}!${nm.cell.ref} became the master of shared formula ${si} (was ${e.pos.ref}).`);
  }

  // --- Column / row default styles for new cells.
  const colStyles = [];
  const colsEl = kids.find((k) => k.local === "cols");
  if (colsEl && !colsEl.selfClosing) {
    for (const c of X.scanChildren(xml, colsEl.openEnd, colsEl.closeStart)) {
      const a = X.parseAttrs(xml, c);
      const s = X.getAttr(a, "style");
      if (s && s !== "0") colStyles.push({ min: +X.getAttr(a, "min"), max: +X.getAttr(a, "max"), style: s });
    }
  }
  const colStyle = (col) => colStyles.find((c) => col >= c.min && col <= c.max)?.style;

  const byRow = new Map();
  for (const e of cellEdits) {
    if (!byRow.has(e.pos.row)) byRow.set(e.pos.row, new Map());
    byRow.get(e.pos.row).set(e.pos.col, e);
  }
  for (const ref of fRewrites.keys()) {
    const p = parseCellRef(ref);
    if (!byRow.has(p.row)) byRow.set(p.row, new Map());
  }
  const formulasRemoved = [];
  const cName = X.withPrefix(sd.name, "c");
  const ctxFor = (r) => ({ cName, sheet, wb, sst, formulaAt, fRewrites, formulasRemoved, rowStyle: r });

  // --- Rebuild sheetData: untouched rows are copied verbatim.
  const rowNum = (r) => parseInt(X.getAttr(X.parseAttrs(xml, r), "r"), 10);
  const targetRows = [...byRow.keys()].sort((a, b) => a - b);
  let body = "";
  let pos = sd.selfClosing ? null : sd.openEnd;
  let ti = 0;
  const rowName = X.withPrefix(sd.name, "row");
  for (const r of rows) {
    const n = rowNum(r);
    while (ti < targetRows.length && targetRows[ti] < n) {
      const tr = targetRows[ti++];
      body += newRow(tr);
    }
    body += xml.slice(pos, r.start);
    if (ti < targetRows.length && targetRows[ti] === n) {
      ti++;
      body += rewriteRow(r, n);
    } else body += xml.slice(r.start, r.end);
    pos = r.end;
  }
  const tail = sd.selfClosing ? "" : xml.slice(pos, sd.closeStart);
  let trailing = "";
  while (ti < targetRows.length) trailing += newRow(targetRows[ti++]);

  function newRow(n) {
    const cells = [...byRow.get(n).entries()].sort((a, b) => a[0] - b[0]);
    const inner = cells.map(([col, e]) => buildCell(null, e, ctxFor(null), colStyle(col))).join("");
    return inner ? `<${rowName} r="${n}">${inner}</${rowName}>` : "";
  }

  function rewriteRow(r, n) {
    const edits = byRow.get(n);
    let attrs = X.parseAttrs(xml, r);
    const rowS = /^(1|true)$/.test(X.getAttr(attrs, "customFormat") || "") ? X.getAttr(attrs, "s") : undefined;
    const children = r.selfClosing ? [] : X.scanChildren(xml, r.openEnd, r.closeStart);
    // Explicit cell positions (add r to cells that rely on implicit positions).
    let prevCol = 0;
    const cells = children.map((c) => {
      if (c.local !== "c") return { el: c };
      const a = X.parseAttrs(xml, c);
      const cr = X.getAttr(a, "r");
      const col = cr ? parseCellRef(cr).col : prevCol + 1;
      prevCol = col;
      return { el: c, col, attrs: a, implicit: !cr };
    });
    const anyImplicit = cells.some((c) => c.implicit);
    const pending = [...edits.entries()].filter(([col]) => !cells.some((c) => c.col === col)).sort((a, b) => a[0] - b[0]);
    const maxNew = pending.length ? pending[pending.length - 1][0] : 0;
    const minNew = pending.length ? pending[0][0] : 0;
    // "spans" is an optional hint; drop it if new cells fall outside it.
    const spans = X.getAttr(attrs, "spans");
    if (spans && pending.length) {
      const ranges = spans.split(/\s+/).map((s) => s.split(":").map(Number));
      if (!ranges.some(([a, b]) => minNew >= a && maxNew <= b)) attrs = X.removeAttrs(attrs, ["spans"]);
    }
    let out = X.startTag(r.name, attrs);
    let p = r.selfClosing ? null : r.openEnd;
    let pi = 0;
    const emitNew = () => {
      const [col, e] = pending[pi++];
      return buildCell(null, e, ctxFor(rowS), rowS ?? colStyle(col));
    };
    for (const c of cells) {
      if (c.col === undefined) {
        // non-cell child (extLst): flush remaining new cells before it
        while (pi < pending.length) out += emitNew();
        out += xml.slice(p, c.el.end);
        p = c.el.end;
        continue;
      }
      while (pi < pending.length && pending[pi][0] < c.col) out += emitNew();
      out += xml.slice(p, c.el.start);
      const ref = `${numToCol(c.col)}${n}`;
      const e = edits.get(c.col);
      if (e) out += buildCell({ xml, el: c.el, attrs: c.implicit ? [{ name: "r", value: ref, quote: '"' }, ...c.attrs] : c.attrs }, e, ctxFor(rowS));
      else if (fRewrites.has(ref)) out += rewriteF(c, ref);
      else if (anyImplicit && c.implicit) out += X.startTag(c.el.name, [{ name: "r", value: ref, quote: '"' }, ...c.attrs], c.el.selfClosing) + xml.slice(c.el.openEnd, c.el.end);
      else out += xml.slice(c.el.start, c.el.end);
      p = c.el.end;
    }
    while (pi < pending.length) out += emitNew();
    out += r.selfClosing ? "" : xml.slice(p, r.closeStart);
    return out + `</${r.name}>`;
  }

  function rewriteF(c, ref) {
    const f = X.scanChildren(xml, c.el.openEnd, c.el.closeStart).find((k) => k.local === "f");
    return xml.slice(c.el.start, f.start) + fRewrites.get(ref) + xml.slice(f.end, c.el.end);
  }

  const newSheetData = sd.selfClosing
    ? `<${sd.name}>${body}${trailing}</${sd.name}>`
    : xml.slice(sd.start, sd.openEnd) + body + tail + trailing + `</${sd.name}>`;
  let result = xml.slice(0, sd.start) + newSheetData + xml.slice(sd.end);

  // --- <dimension ref>: grow to include written cells (optional element).
  const dim = kids.find((k) => k.local === "dimension");
  const written = cellEdits.filter((e) => e.kind !== "clear");
  if (dim && written.length && dim.end <= sd.start) {
    const a = X.parseAttrs(xml, dim);
    const cur = X.getAttr(a, "ref");
    let g = { r1: Infinity, c1: Infinity, r2: 0, c2: 0 };
    try {
      if (cur) g = parseRangeRef(cur);
    } catch {}
    const empty = cur === "A1" && rows.length === 0;
    if (empty) g = { r1: Infinity, c1: Infinity, r2: 0, c2: 0 };
    for (const e of written) g = { r1: Math.min(g.r1, e.pos.row), c1: Math.min(g.c1, e.pos.col), r2: Math.max(g.r2, e.pos.row), c2: Math.max(g.c2, e.pos.col) };
    const ref = g.r1 === g.r2 && g.c1 === g.c2 ? `${numToCol(g.c1)}${g.r1}` : `${numToCol(g.c1)}${g.r1}:${numToCol(g.c2)}${g.r2}`;
    if (ref !== cur) result = xml.slice(0, dim.start) + X.startTag(dim.name, X.setAttr(a, "ref", ref), dim.selfClosing) + xml.slice(dim.openEnd, sd.start) + newSheetData + xml.slice(sd.end);
  }
  X.assertWellFormed(result, sheet.path);
  pkg.set(sheet.path, result);
  return { notes, formulasRemoved };
}

/**
 * Build the XML for one cell. `existing` = { xml, el, attrs } or null for a new cell.
 * Keeps r, s and unknown attributes/children (e.g. extLst); drops t/cm/vm, which describe
 * the old value.
 */
function buildCell(existing, edit, ctx, inheritedStyle) {
  const { cName, wb, sst, formulaAt, formulasRemoved } = ctx;
  let attrs = existing ? [...existing.attrs] : [{ name: "r", value: edit.pos.ref, quote: '"' }];
  if (!existing && inheritedStyle && inheritedStyle !== "0") attrs.push({ name: "s", value: inheritedStyle, quote: '"' });
  attrs = X.removeAttrs(attrs, ["t", "cm", "vm"]);
  const name = existing ? existing.el.name : cName;
  let extra = "";
  if (existing && !existing.el.selfClosing) {
    for (const k of X.scanChildren(existing.xml, existing.el.openEnd, existing.el.closeStart)) {
      if (!["f", "v", "is"].includes(k.local)) extra += existing.xml.slice(k.start, k.end);
    }
  }
  const child = (local) => X.withPrefix(name, local);
  const hadFormula = formulaAt.has(edit.pos.ref);
  let inner = "";
  switch (edit.kind) {
    case "number":
      inner = `<${child("v")}>${numberText(edit.value)}</${child("v")}>`;
      break;
    case "date":
      inner = `<${child("v")}>${numberText(excelSerial(edit.value, wb.date1904))}</${child("v")}>`;
      break;
    case "boolean":
      attrs = placeT(attrs, "b");
      inner = `<${child("v")}>${edit.value ? 1 : 0}</${child("v")}>`;
      break;
    case "string": {
      if (sst) {
        const { index } = sst.indexOf(edit.value);
        attrs = placeT(attrs, "s");
        inner = `<${child("v")}>${index}</${child("v")}>`;
      } else {
        attrs = placeT(attrs, "inlineStr");
        const enc = X.encodeXstring(edit.value);
        inner = `<${child("is")}><${child("t")}${/^\s|\s$/.test(enc) ? ' xml:space="preserve"' : ""}>${X.escapeText(enc)}</${child("t")}></${child("is")}>`;
      }
      break;
    }
    case "formula":
      inner = `<${child("f")}>${X.escapeText(edit.value)}</${child("f")}>`;
      break;
    case "clear":
      inner = "";
      break;
  }
  if (hadFormula && edit.kind !== "formula") formulasRemoved.push(edit.pos.ref);
  if (edit.kind === "clear" && !extra && attrs.every((a) => a.name === "r")) return ""; // nothing left to keep
  const body = inner + extra;
  return body ? `${X.startTag(name, attrs)}${body}</${name}>` : X.startTag(name, attrs, true);
}

/** Insert t after s (or r), matching Excel's attribute order. */
function placeT(attrs, t) {
  const a = X.removeAttrs(attrs, ["t"]);
  const after = a.findIndex((x) => x.name === "s") >= 0 ? "s" : "r";
  return X.setAttr(a, "t", t, after);
}

// ---------- calcChain & calcPr ----------

function pruneCalcChain(pkg, wb, removals) {
  if (!wb.calcChainPath || !removals.length || !pkg.has(wb.calcChainPath)) return 0;
  const drop = new Set(removals.map((r) => `${r.sheetId}|${r.ref}`));
  const xml = pkg.text(wb.calcChainPath);
  const root = X.rootElement(xml);
  const entries = root.selfClosing ? [] : X.scanChildren(xml, root.openEnd, root.closeStart).filter((e) => e.local === "c");
  let effI = null;
  let lastKeptI = null;
  let removed = 0;
  let out = xml.slice(0, root.openEnd);
  let pos = root.openEnd;
  for (const e of entries) {
    const a = X.parseAttrs(xml, e);
    const i = X.getAttr(a, "i");
    if (i !== undefined) effI = i;
    const ref = parseCellRef(X.getAttr(a, "r")).ref;
    if (drop.has(`${effI}|${ref}`)) {
      removed++;
      pos = e.end;
      continue;
    }
    out += xml.slice(pos, e.start);
    // The sheet id is inherited from the previous entry; make it explicit if the entry it
    // inherited from was removed.
    if (i === undefined && effI !== lastKeptI) out += X.startTag(e.name, [...a, { name: "i", value: effI, quote: '"' }], e.selfClosing) + xml.slice(e.openEnd, e.end);
    else out += xml.slice(e.start, e.end);
    lastKeptI = effI;
    pos = e.end;
  }
  if (!removed) return 0;
  if (removed === entries.length) {
    // An empty calcChain is invalid: remove the part, its relationship and content type.
    pkg.remove(wb.calcChainPath);
    const fresh = readRels(pkg, wb.path); // re-read: add_sheet may have changed this part
    const rel = fresh.rels.find((r) => r.type === REL.calcChain)?.el;
    const relsXml = pkg.text(wb.rels.path);
    if (rel) pkg.set(wb.rels.path, relsXml.slice(0, rel.start) + relsXml.slice(rel.end));
    const ct = pkg.text("[Content_Types].xml");
    const ctRoot = X.rootElement(ct);
    const ov = X.scanChildren(ct, ctRoot.openEnd, ctRoot.closeStart).find(
      (k) => k.local === "Override" && (X.getAttr(X.parseAttrs(ct, k), "PartName") || "").toLowerCase() === "/" + wb.calcChainPath.toLowerCase()
    );
    if (ov) pkg.set("[Content_Types].xml", ct.slice(0, ov.start) + ct.slice(ov.end));
    return removed;
  }
  pkg.set(wb.calcChainPath, out + xml.slice(pos));
  return removed;
}

function ensureFullCalcOnLoad(pkg, wb) {
  const xml = pkg.text(wb.path);
  const root = X.rootElement(xml);
  const kids = X.scanChildren(xml, root.openEnd, root.closeStart);
  const calcPr = kids.find((k) => k.local === "calcPr");
  if (calcPr) {
    const a = X.parseAttrs(xml, calcPr);
    if (/^(1|true)$/.test(X.getAttr(a, "fullCalcOnLoad") || "")) return false;
    pkg.set(wb.path, xml.slice(0, calcPr.start) + X.startTag(calcPr.name, X.setAttr(a, "fullCalcOnLoad", "1"), calcPr.selfClosing) + xml.slice(calcPr.openEnd));
    return true;
  }
  const before = kids.find((k) => WORKBOOK_AFTER_CALCPR.includes(k.local));
  const at = before ? before.start : root.closeStart;
  pkg.set(wb.path, xml.slice(0, at) + `<${X.withPrefix(root.name, "calcPr")} fullCalcOnLoad="1"/>` + xml.slice(at));
  return true;
}

// ---------- edit validation ----------

function normalizeEdits(edits) {
  if (!Array.isArray(edits) || edits.length === 0) throw new Error("edits must be a non-empty array.");
  if (edits.length > MAX_EDITS) throw new Error(`Too many edits (${edits.length}); the limit is ${MAX_EDITS} per call.`);
  const adds = [];
  const cells = [];
  edits.forEach((e, i) => {
    const where = `edits[${i}]`;
    if (!e || typeof e !== "object" || Array.isArray(e)) throw new Error(`${where} must be an object.`);
    if ("add_sheet" in e) {
      if (Object.keys(e).some((k) => k !== "add_sheet")) throw new Error(`${where}: add_sheet cannot be combined with other fields.`);
      adds.push(e.add_sheet);
      return;
    }
    const keys = ["value", "formula", "clear"].filter((k) => k in e && e[k] !== undefined);
    if (keys.length !== 1) throw new Error(`${where} must have exactly one of "value", "formula" or "clear" (plus "sheet" and "cell").`);
    if (typeof e.sheet !== "string" || !e.sheet) throw new Error(`${where}: "sheet" is required.`);
    const pos = parseCellRef(e.cell);
    let kind;
    let value;
    if (keys[0] === "clear") {
      if (e.clear !== true) throw new Error(`${where}: "clear" must be true.`);
      kind = "clear";
    } else if (keys[0] === "formula") {
      kind = "formula";
      value = normalizeFormula(e.formula);
    } else {
      const v = e.value;
      if (v === null) kind = "clear";
      else if (typeof v === "number") {
        if (!Number.isFinite(v)) throw new Error(`${where}: numbers must be finite.`);
        kind = "number";
        value = v;
      } else if (typeof v === "boolean") {
        kind = "boolean";
        value = v;
      } else if (typeof v === "string") {
        if (v.length > 32767) throw new Error(`${where}: text is longer than Excel's 32767-character cell limit.`);
        X.encodeXstring(v); // validates surrogates
        kind = "string";
        value = v;
      } else if (v && typeof v === "object" && typeof v.date === "string" && Object.keys(v).length === 1) {
        excelSerial(v.date, false); // validate format early
        kind = "date";
        value = v.date;
      } else throw new Error(`${where}: unsupported value ${JSON.stringify(v)}. Use a number, string, boolean, null or {"date": "YYYY-MM-DD"}.`);
    }
    cells.push({ index: i, sheet: e.sheet, pos, kind, value });
  });
  return { adds, cells };
}

// ---------- the transaction ----------

/** Apply edits to a package buffer in memory. Returns { buffer, report }. */
export function applyEdits(srcBuf, edits) {
  const { adds, cells } = normalizeEdits(edits);
  const pkg = new Package(srcBuf);
  const wb = loadWorkbook(pkg);
  const notes = [];

  const addedSheets = adds.map((name) => addSheet(pkg, wb, name).name);

  // Group cell edits per sheet; later edits to the same cell win.
  const perSheet = new Map();
  for (const c of cells) {
    const sheet = findSheet(wb, c.sheet);
    c.sheetName = sheet.name;
    if (!perSheet.has(sheet)) perSheet.set(sheet, new Map());
    perSheet.get(sheet).set(c.pos.ref, c);
  }
  const sst = wb.sstPath && pkg.has(wb.sstPath) ? new SharedStrings(pkg, wb.sstPath) : null;
  const removals = [];
  for (const [sheet, map] of perSheet) {
    const res = editWorksheet(pkg, wb, sheet, [...map.values()], sst);
    notes.push(...res.notes);
    for (const ref of res.formulasRemoved) removals.push({ sheetId: sheet.sheetId, ref });
  }
  const sstChanged = sst ? sst.finalize() : false;
  if (sstChanged) notes.push(`Added ${sst.pending.length} new shared string(s).`);
  const pruned = pruneCalcChain(pkg, wb, removals);
  if (pruned) notes.push(`Removed ${pruned} stale calculation-chain entr${pruned === 1 ? "y" : "ies"} for cells that no longer contain formulas.`);
  if (cells.length) ensureFullCalcOnLoad(pkg, wb);

  // Well-formedness of every modified/added XML part.
  for (const name of [...pkg.changed].map((k) => pkg.byName.get(k)?.name).filter(Boolean).concat(pkg.additions.map((a) => a.name))) {
    if (/\.(xml|rels)$/i.test(name) && !pkg.deleted.has(name.toLowerCase())) X.assertWellFormed(pkg.text(name), name);
  }
  const buffer = pkg.build();
  return {
    buffer,
    cells,
    addedSheets,
    notes,
    modified: pkg.modifiedPartNames(),
    added: pkg.additions.map((a) => a.name),
    removed: [...pkg.deleted].map((k) => pkg.byName.get(k)?.name || k),
    untouched: pkg.zip.entries.filter((e) => !pkg.changed.has(e.name.toLowerCase()) && !pkg.deleted.has(e.name.toLowerCase())).length,
    totalParts: pkg.zip.entries.length,
  };
}

/** Independent checks on the bytes about to be committed. */
async function validateOutput(outBuf, srcBuf, result) {
  // 1. Our own reader: structure + CRC of every re-encoded part.
  const out = readZip(outBuf);
  const src = readZip(srcBuf);
  const outBy = new Map(out.entries.map((e) => [e.name.toLowerCase(), e]));
  const touched = new Set([...result.modified, ...result.added, ...result.removed].map((n) => n.toLowerCase()));
  for (const e of out.entries) if (touched.has(e.name.toLowerCase())) readEntry(e);
  // 2. Every untouched entry must be byte-identical (raw local record + data).
  for (const e of src.entries) {
    const k = e.name.toLowerCase();
    if (touched.has(k)) continue;
    const o = outBy.get(k);
    if (!o || !o.raw.equals(e.raw)) throw new Error(`Internal check failed: untouched part ${e.name} changed.`);
  }
  // 3. An independent ZIP implementation with CRC checks on all entries.
  const jz = await JSZip.loadAsync(outBuf, { checkCRC32: true });
  await Promise.all(Object.values(jz.files).filter((f) => !f.dir).map((f) => f.async("uint8array")));
  // 4. Semantic read-back with ExcelJS (read-only use). If ExcelJS cannot read the
  //    source either, skip this step rather than block the edit.
  const readBook = async (buf) => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    return wb;
  };
  let book;
  try {
    book = await readBook(outBuf);
  } catch (err) {
    try {
      await readBook(srcBuf);
    } catch {
      return "ExcelJS cannot read this workbook type, so the semantic read-back check was skipped (ZIP and XML checks passed).";
    }
    throw new Error(`Validation failed: the edited workbook could not be read back (${err.message}). Nothing was written.`);
  }
  for (const c of result.cells) {
    const ws = book.worksheets.find((w) => w.name.toLowerCase() === c.sheetName.toLowerCase());
    if (!ws) throw new Error(`Validation failed: sheet ${c.sheetName} missing after edit.`);
    const v = ws.getCell(c.pos.ref).value;
    const ok =
      c.kind === "clear"
        ? v === null || v === undefined || v === ""
        : c.kind === "formula"
          ? v && typeof v === "object" && (v.formula || "").replace(/^=/, "") === c.value
          : c.kind === "string"
            ? (typeof v === "string" ? v : v?.richText?.map((r) => r.text).join("")) === c.value
            : c.kind === "boolean"
              ? v === c.value
              : c.kind === "number"
                ? v === c.value || (v instanceof Date && typeof c.value === "number")
                : c.kind === "date"
                  ? v instanceof Date || typeof v === "number"
                  : false;
    if (!ok) throw new Error(`Validation failed: ${c.sheetName}!${c.pos.ref} reads back as ${JSON.stringify(v)} after the edit. Nothing was written.`);
  }
  return null;
}

// ---------- locking ----------

const inProcessLocks = new Map();
const LOCK_STALE_MS = 60_000;

function lockKey(p) {
  return process.platform === "win32" || process.platform === "darwin" ? p.toLowerCase() : p;
}

async function acquireInProcess(key) {
  const prev = inProcessLocks.get(key) || Promise.resolve();
  let release;
  const mine = new Promise((r) => (release = r));
  const tail = prev.then(() => mine);
  inProcessLocks.set(key, tail);
  await prev;
  return () => {
    release();
    if (inProcessLocks.get(key) === tail) inProcessLocks.delete(key);
  };
}

function pidAlive(pid) {
  try {
    process.kill(pid, 0);
    return true;
  } catch (e) {
    return e.code === "EPERM";
  }
}

/** Cross-process lock file next to the target; skipped if the folder is not writable. */
async function acquireLockFile(file) {
  const lockPath = path.join(path.dirname(file), `.~localdocs-lock.${path.basename(file)}`);
  const timeout = Number(process.env.LOCAL_DOCUMENTS_LOCK_TIMEOUT_MS || 15_000);
  const started = Date.now();
  const me = JSON.stringify({ pid: process.pid, host: os.hostname(), time: Date.now() });
  for (;;) {
    try {
      fs.writeFileSync(lockPath, me, { flag: "wx" });
      return () => {
        try {
          if (fs.readFileSync(lockPath, "utf8") === me) fs.unlinkSync(lockPath);
        } catch {}
      };
    } catch (err) {
      if (["EACCES", "EPERM", "EROFS"].includes(err.code)) return () => {};
      if (err.code !== "EEXIST") throw err;
    }
    let stale = false;
    try {
      const info = JSON.parse(fs.readFileSync(lockPath, "utf8"));
      stale = Date.now() - info.time > LOCK_STALE_MS || (info.host === os.hostname() && !pidAlive(info.pid));
    } catch {
      stale = true;
    }
    if (stale) {
      try {
        fs.unlinkSync(lockPath);
      } catch {}
      continue;
    }
    if (Date.now() - started > timeout) throw new Error(`${path.basename(file)} is being edited by another process (lock file ${lockPath}). Try again shortly.`);
    await new Promise((r) => setTimeout(r, 100));
  }
}

async function withFileLocks(files, fn) {
  const keys = [...new Set(files.map(lockKey))].sort();
  const releases = [];
  try {
    for (const k of keys) {
      releases.push(await acquireInProcess(k));
      releases.push(await acquireLockFile(files.find((f) => lockKey(f) === k)));
    }
    return await fn();
  } finally {
    for (const r of releases.reverse()) r();
  }
}

// ---------- output planning & commit ----------

function defaultOutputName(source) {
  const dir = path.dirname(source);
  const ext = path.extname(source);
  const stem = path.basename(source, ext);
  for (let i = 1; i < 1000; i++) {
    const name = `${stem} (edited${i === 1 ? "" : " " + i})${ext}`;
    if (!fs.existsSync(path.join(dir, name))) return name;
  }
  throw new Error("Could not find a free output file name; pass output_path explicitly.");
}

function planOutput(source, outputPath, overwrite) {
  const ext = path.extname(source.realPath).toLowerCase();
  const baseDir = path.dirname(source.realPath);
  if (outputPath === undefined || outputPath === null || outputPath === "") {
    if (overwrite) return { target: source.realPath, overwritesSource: true, replacesExisting: true };
    const out = resolveAllowedOutput(defaultOutputName(source.realPath), { extensions: [ext], baseDir, overwrite: false });
    return { target: out.path, overwritesSource: false, replacesExisting: false, isDefault: true };
  }
  const out = resolveAllowedOutput(outputPath, { extensions: [ext], baseDir, overwrite: Boolean(overwrite) });
  return { target: out.path, overwritesSource: samePath(out.path, source.realPath), replacesExisting: out.exists };
}

/** Write the new package to a temp file in the destination folder. */
function commit(buffer, plan) {
  const dir = path.dirname(plan.target);
  const tmp = path.join(dir, `.~${path.basename(plan.target)}.${process.pid}.${crypto.randomBytes(4).toString("hex")}.tmp`);
  const fd = fs.openSync(tmp, "wx");
  try {
    fs.writeSync(fd, buffer);
    fs.fsyncSync(fd);
  } finally {
    fs.closeSync(fd);
  }
  return tmp;
}

function finalize(tmp, plan, expectedHash) {
  try {
    assertOutputStillAllowed(plan.target);
    if (plan.replacesExisting) {
      // Compare-and-swap: refuse if the file changed since we read it.
      let current;
      try {
        current = sha256(fs.readFileSync(plan.target));
      } catch (e) {
        throw new Error(`${plan.target} disappeared while editing; nothing was written.`);
      }
      if (current !== expectedHash) throw new Error(`${plan.target} was changed by someone else while editing; nothing was written. Re-run the edit.`);
      try {
        fs.chmodSync(tmp, fs.statSync(plan.target).mode & 0o7777);
      } catch {}
      fs.renameSync(tmp, plan.target);
    } else {
      // Create without clobbering: a hard link fails if the target exists.
      try {
        fs.linkSync(tmp, plan.target);
        fs.unlinkSync(tmp);
      } catch (e) {
        if (e.code === "EEXIST") throw Object.assign(new Error(`Output file appeared while writing: ${plan.target}`), { code: "LD_EXISTS" });
        if (fs.existsSync(plan.target)) throw Object.assign(new Error(`Output file appeared while writing: ${plan.target}`), { code: "LD_EXISTS" });
        fs.renameSync(tmp, plan.target); // file systems without hard links
      }
    }
  } catch (err) {
    try {
      fs.unlinkSync(tmp);
    } catch {}
    if (err?.code === "EBUSY" || err?.code === "EPERM") throw new Error(`Could not write ${plan.target}: the file is locked (is it open in Excel?). ${err.message}`);
    throw err;
  }
}

function describe(result, plan, source, extraNote) {
  const lines = [];
  const bySheet = new Map();
  for (const c of result.cells) bySheet.set(c.sheetName, (bySheet.get(c.sheetName) || 0) + 1);
  if (result.addedSheets.length) lines.push(`Added sheet(s): ${result.addedSheets.map((s) => JSON.stringify(s)).join(", ")}.`);
  if (result.cells.length) lines.push(`Applied ${result.cells.length} cell edit(s): ${[...bySheet].map(([s, n]) => `${n} on "${s}"`).join(", ")}.`);
  lines.push(plan.overwritesSource ? `Saved over the source file: ${plan.target}` : `Saved to new file: ${plan.target}`);
  if (!plan.overwritesSource) lines.push(`The source file was not modified: ${source.realPath}`);
  lines.push(`Package parts: ${result.untouched} of ${result.totalParts} copied byte-for-byte; modified: ${result.modified.join(", ") || "none"}${result.added.length ? `; added: ${result.added.join(", ")}` : ""}${result.removed.length ? `; removed: ${result.removed.join(", ")}` : ""}.`);
  lines.push(...result.notes);
  if (result.cells.length) lines.push("Formulas are not recalculated by this tool; the workbook is flagged so Excel recalculates everything when it is opened (cached values of other formulas may be stale until then).");
  if (extraNote) lines.push(extraNote);
  return lines.join("\n");
}

export async function xlsxEdit(inputPath, edits, { outputPath, overwrite = false } = {}) {
  normalizeEdits(edits); // fail fast on malformed input
  const source = resolveAllowedFile(inputPath, EDIT_EXT);
  let plan = planOutput(source, outputPath, overwrite);
  return withFileLocks([source.realPath, ...(samePath(plan.target, source.realPath) ? [] : [plan.target])], async () => {
    for (let attempt = 0; ; attempt++) {
      const srcBuf = fs.readFileSync(source.realPath);
      const srcHash = sha256(srcBuf);
      const expectedHash = plan.replacesExisting ? (plan.overwritesSource ? srcHash : sha256(fs.readFileSync(plan.target))) : null;
      const result = applyEdits(srcBuf, edits);
      const tmp = commit(result.buffer, plan);
      let note;
      try {
        note = await validateOutput(fs.readFileSync(tmp), srcBuf, result); // validate what is on disk
      } catch (err) {
        try {
          fs.unlinkSync(tmp);
        } catch {}
        throw err;
      }
      try {
        finalize(tmp, plan, expectedHash);
      } catch (err) {
        if (err.code === "LD_EXISTS" && plan.isDefault && attempt < 3) {
          plan = planOutput(source, undefined, false);
          continue;
        }
        throw err;
      }
      return describe(result, plan, source, note);
    }
  });
}
