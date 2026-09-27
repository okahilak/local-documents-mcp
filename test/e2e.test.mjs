// End-to-end tests: spawns the server over stdio and calls tools through an MCP client.
// The whole suite runs twice: with a legacy (2025-era, SDK 1.x) client doing the plain
// initialize handshake, and with a Claude-Desktop-like SDK 2.x client that probes with
// server/discover in place before negotiating.
// SERVER_ENTRY env var can point at an unpacked bundle's server/index.js.
import { describe, test, before, after } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { connectLikeClaude } from "./claude-like-client.mjs";
import { PNG } from "pngjs";

const here = path.dirname(fileURLToPath(import.meta.url));
const fx = path.join(here, "fixtures");
const allowed = path.join(fx, "allowed");
const entry = path.resolve(process.env.SERVER_ENTRY || path.join(here, "..", "server", "index.js"));
const serverArgs = [entry, allowed, "${user_config.allowed_directories}", path.join(fx, "does-not-exist")];

const clients = {
  "legacy client (SDK 1.x, initialize only)": async () => {
    const transport = new StdioClientTransport({ command: process.execPath, args: serverArgs, stderr: "pipe" });
    let err = "";
    transport.stderr.on("data", (d) => (err += d));
    const c = new Client({ name: "e2e", version: "1.0.0" });
    await c.connect(transport);
    return { client: c, stderr: () => err };
  },
  "Claude-like client (SDK 2.x, in-place server/discover probe)": async () => connectLikeClaude(process.execPath, serverArgs),
};

for (const [label, connect] of Object.entries(clients)) describe(label, () => {
let client;
let getStderr;
before(async () => {
  ({ client, stderr: getStderr } = await connect());
});
after(async () => client?.close());

const call = (name, args) => client.callTool({ name, arguments: args });
const text = (r) => r.content.filter((c) => c.type === "text").map((c) => c.text).join("\n");

test("lists the ten read-only tools and the xlsx_edit tool", async () => {
  const { tools } = await client.listTools();
  const editing = ["xlsx_edit"];
  assert.deepEqual(tools.map((t) => t.name).sort(), [
    "docx_info", "docx_read_tables", "docx_read_text", "docx_search",
    "pdf_info", "pdf_read_text", "pdf_render_page", "xlsx_edit", "xlsx_info", "xlsx_read_range", "xlsx_search",
  ]);
  for (const t of tools) assert.equal(t.annotations.readOnlyHint, !editing.includes(t.name), t.name);
});

test("pdf_info", async () => {
  const r = await call("pdf_info", { path: path.join(allowed, "report.pdf") });
  assert.ok(!r.isError, text(r));
  const info = JSON.parse(text(r));
  assert.equal(info.pages, 3);
  assert.equal(info.title, "Quarterly Report");
  assert.equal(info.author, "Test Author");
  assert.equal(info.text_layer, "present");
  const s = JSON.parse(text(await call("pdf_info", { path: "sub/scanned.pdf" })));
  assert.match(s.text_layer, /absent/);
});

test("pdf_read_text all pages and page spec", async () => {
  let r = await call("pdf_read_text", { path: path.join(allowed, "report.pdf") });
  let t = text(r);
  assert.ok(!r.isError, t);
  assert.match(t, /--- Page 1 ---\nPage 1 heading/);
  assert.match(t, /Revenue for region 3 was 3000 units\./);
  assert.match(t, /units\.\nSecond line of body text\./);
  r = await call("pdf_read_text", { path: "report.pdf", pages: "2-3" });
  t = text(r);
  assert.doesNotMatch(t, /Page 1 heading/);
  assert.match(t, /Page 2 heading[\s\S]*Page 3 heading/);
  r = await call("pdf_read_text", { path: "report.pdf", pages: 2 });
  assert.match(text(r), /region 2/);
  r = await call("pdf_read_text", { path: "report.pdf", pages: "9" });
  assert.ok(r.isError);
  r = await call("pdf_read_text", { path: "sub/scanned.pdf" });
  assert.match(text(r), /no text layer/);
});

test("pdf_render_page returns a real image", async () => {
  const r = await call("pdf_render_page", { path: "sub/scanned.pdf", page: 1, format: "png" });
  assert.ok(!r.isError, text(r));
  const img = r.content.find((c) => c.type === "image");
  assert.equal(img.mimeType, "image/png");
  const png = PNG.sync.read(Buffer.from(img.data, "base64"));
  assert.equal(png.width, 1224); // 612pt at 144 dpi
  assert.equal(png.height, 1584);
  // centre of the red box should be red, top-left corner white
  const px = (x, y) => Array.from(png.data.subarray((y * png.width + x) * 4, (y * png.width + x) * 4 + 3));
  const [r0, g0, b0] = px(612, 792);
  assert.ok(r0 > 150 && g0 < 80 && b0 < 80, `expected red, got ${[r0, g0, b0]}`);
  assert.deepEqual(px(5, 5), [255, 255, 255]);
  fs.writeFileSync(path.join(fx, "render-test.png"), Buffer.from(img.data, "base64"));

  const j = await call("pdf_render_page", { path: "report.pdf", page: 2, format: "jpeg", dpi: 72 });
  const jimg = j.content.find((c) => c.type === "image");
  assert.equal(jimg.mimeType, "image/jpeg");
  assert.equal(Buffer.from(jimg.data, "base64").subarray(0, 2).toString("hex"), "ffd8");
  const bad = await call("pdf_render_page", { path: "report.pdf", page: 4 });
  assert.ok(bad.isError);
});

test("xlsx_info", async () => {
  const r = await call("xlsx_info", { path: path.join(allowed, "book.xlsx") });
  assert.ok(!r.isError, text(r));
  const info = JSON.parse(text(r));
  assert.equal(info.sheet_count, 2);
  assert.equal(info.sheets[0].name, "Sales");
  assert.equal(info.sheets[1].state, "hidden");
  assert.equal(info.sheets[1].used_range, "A1:J40");
  assert.ok(info.defined_names.some((d) => d.name === "Quarters"));
});

test("xlsx_read_range values, formulas, dates, merges", async () => {
  const r = await call("xlsx_read_range", { path: "book.xlsx", sheet: "sales", range: "A1:E6", include_formulas: true });
  const t = text(r);
  assert.ok(!r.isError, t);
  assert.match(t, /\tA\tB\tC\tD\tE/);
  assert.match(t, /\n2\tNorth\t100\t150\t250\t2026-01-15/);
  assert.match(t, /\n3\tSouth\t80\.5\t20\t100\.5\t2026-02-01/);
  assert.match(t, /\n4\tMulti\\nline\tTRUE\trich text\t#DIV\/0!\t/);
  assert.match(t, /\n6\tMerged banner\t\t\t\t/);
  assert.match(t, /D2: =B2\+C2\s+→\s+250/);
  const cols = await call("xlsx_read_range", { path: "book.xlsx", sheet: "1", range: "B:B" });
  assert.match(text(cols), /\n3\t80\.5/);
  const single = await call("xlsx_read_range", { path: "book.xlsx", sheet: "Hidden Notes", range: "$J$40" });
  assert.match(text(single), /\n40\tExample link/);
  const badSheet = await call("xlsx_read_range", { path: "book.xlsx", sheet: "Nope", range: "A1" });
  assert.ok(badSheet.isError);
  assert.match(text(badSheet), /Available sheets/);
  const badRange = await call("xlsx_read_range", { path: "book.xlsx", sheet: "Sales", range: "A1:ZZZZ9" });
  assert.ok(badRange.isError);
});

test("xlsx_search", async () => {
  let r = await call("xlsx_search", { path: "book.xlsx", query: "NEEDLE" });
  assert.match(text(r), /1 match/);
  assert.match(text(r), /Hidden Notes!B3\tneedle in hidden sheet/);
  r = await call("xlsx_search", { path: "book.xlsx", query: "B2+", include_formulas: true });
  assert.match(text(r), /Sales!D2/);
  r = await call("xlsx_search", { path: "book.xlsx", query: "north", match_case: true });
  assert.match(text(r), /0 matches/);
});

test("docx_info: properties, outline, counts", async () => {
  const r = await call("docx_info", { path: path.join(allowed, "memo.docx") });
  assert.ok(!r.isError, text(r));
  const info = JSON.parse(text(r));
  assert.equal(info.title, "Project Memo");
  assert.equal(info.author, "Jane Tester");
  assert.equal(info.last_modified_by, "Reviewer");
  assert.equal(info.keywords, "memo, fixture");
  assert.deepEqual(
    { headings: info.counts.headings, tables: info.counts.tables, list_items: info.counts.list_items, images: info.counts.images, footnotes: info.counts.footnotes, comments: info.counts.comments },
    { headings: 6, tables: 2, list_items: 7, images: 1, footnotes: 1, comments: 1 }
  );
  assert.deepEqual(info.outline, ["H1 (Title): Project Memo", "H1: Introduction", "  H2: Goals", "  H2: Steps", "H1: Figures", "H1: Conclusion"]);
  assert.equal(info.tables[0].merged_regions, 2);
  assert.equal(info.tables[0].header_rows, 1);
  assert.equal(info.has_page_headers, true);
});

test("docx_read_text: order, headings, lists, tables, notes", async () => {
  const r = await call("docx_read_text", { path: "memo.docx" });
  const t = text(r);
  assert.ok(!r.isError, t);
  // Reading order is preserved.
  const order = ["# Project Memo", "# Introduction", "## Goals", "## Steps", "# Figures", "**Table 1**", "Text between tables.", "**Table 2**", "# Conclusion"];
  let last = -1;
  for (const marker of order) {
    const i = t.indexOf(marker);
    assert.ok(i > last, `"${marker}" out of order`);
    last = i;
  }
  assert.match(t, /The budget was \*\*approved\*\* on Monday\.\[\^1\]/);
  assert.match(t, /\[the plan\]\(https:\/\/example\.com\/plan\)/);
  assert.match(t, /- Reduce costs\n    - Cloud spend\n    - Licences\n- Improve quality/);
  assert.match(t, /1\. Collect data\n    1\. Interview teams\n2\. Write report/);
  assert.match(t, /\| Region \| Q1 \| Q2 \|\n\| --- \| --- \| --- \|\n\| North \| 100 \| 150 \|\n\|  \| 110 \| 160 \|/);
  assert.match(t, /\[image: Revenue chart\]/);
  assert.match(t, /\[\^1\]: Source: internal audit 2026\./);
  assert.match(t, /Rita Reviewer \(2026-09-01\) on "Text between tables\.": Please double-check this figure\./);
  assert.match(t, /Page header: CONFIDENTIAL HEADER/);
  assert.match(t, /café – 東京/);
  const tail = await call("docx_read_text", { path: "memo.docx", offset: t.indexOf("# Conclusion") - t.indexOf("\n\n") - 2 });
  assert.ok(!tail.isError);
});

test("docx_read_tables: grids, merges, single table", async () => {
  let r = await call("docx_read_tables", { path: "memo.docx" });
  let t = text(r);
  assert.ok(!r.isError, t);
  assert.match(t, /## Table 1 of 2 — 4 rows × 3 columns — section: Figures — 1 header row/);
  assert.match(t, /\| Total \(needle\) \|  \| 520 \|/);
  assert.match(t, /Merged cells: A2:A3 \("North"\), A4:B4 \("Total \(needle\)"\)/);
  assert.match(t, /\| Owner \| Finance team \|/);
  r = await call("docx_read_tables", { path: "memo.docx", table: 2 });
  t = text(r);
  assert.doesNotMatch(t, /Table 1 of 2/);
  assert.match(t, /Table 2 of 2/);
  r = await call("docx_read_tables", { path: "memo.docx", table: 3 });
  assert.ok(r.isError);
});

test("docx_search: body, lists, tables, notes, comments, headers", async () => {
  const expect = {
    needle: /\[table 1, cell A4 \(row 4, column 1\) · § Figures\] Total \(needle\)/,
    licences: /\[list item 3 · § Introduction › Goals\] Licences/,
    budget: /\[paragraph 1 · § Introduction\] The budget was approved on Monday\./,
    audit: /\[footnote 1\] Source: internal audit 2026\./,
    "double-check": /\[comment by Rita Reviewer\]/,
    confidential: /\[page header\] CONFIDENTIAL HEADER/,
    "東京": /Ship it/,
  };
  for (const [q, re] of Object.entries(expect)) {
    const r = await call("docx_search", { path: "memo.docx", query: q });
    assert.match(text(r), /1 match/, q);
    assert.match(text(r), re, q);
  }
  const cs = await call("docx_search", { path: "memo.docx", query: "NEEDLE", match_case: true });
  assert.match(text(cs), /0 matches/);
});

test("docx: corrupt file gives a clean error", async () => {
  const r = await call("docx_info", { path: "broken.docx" });
  assert.ok(r.isError);
  assert.match(text(r), /Could not read DOCX/);
});

test("security: rejects paths outside allowed directories", async () => {
  const attempts = [
    path.join(fx, "outside", "secret.pdf"), // absolute outside
    path.join(allowed, "..", "outside", "secret.pdf"), // traversal
    "../outside/secret.pdf", // relative traversal
    "sub/../../outside/secret.pdf",
    path.join(fx, "allowed-evil", "evil.pdf"), // sibling with shared prefix
    path.join(allowed, "link-to-secret.pdf"), // file symlink escape
    path.join(allowed, "link-dir", "secret.pdf"), // directory symlink escape
    "link-dir/secret.pdf",
  ];
  for (const p of attempts) {
    for (const tool of ["pdf_info", "pdf_read_text"]) {
      const r = await call(tool, { path: p });
      assert.ok(r.isError, `${tool} should reject ${p}`);
      assert.match(text(r), /Access denied|outside/, p);
    }
    const r = await call("pdf_render_page", { path: p, page: 1 });
    assert.ok(r.isError, `render should reject ${p}`);
  }
  for (const p of [path.join(fx, "outside", "secret.xlsx"), "link-dir/secret.xlsx", "../outside/secret.xlsx"]) {
    for (const [tool, extra] of [["xlsx_info", {}], ["xlsx_read_range", { sheet: "Sales", range: "A1" }], ["xlsx_search", { query: "a" }]]) {
      const r = await call(tool, { path: p, ...extra });
      assert.ok(r.isError, `${tool} should reject ${p}`);
      assert.match(text(r), /Access denied/);
    }
  }
  for (const p of [path.join(fx, "outside", "secret.docx"), "../outside/secret.docx", "link-dir/secret.docx", "link-to-secret.docx", path.join(allowed, "sub", "..", "..", "outside", "secret.docx")]) {
    for (const [tool, extra] of [["docx_info", {}], ["docx_read_text", {}], ["docx_read_tables", {}], ["docx_search", { query: "a" }]]) {
      const r = await call(tool, { path: p, ...extra });
      assert.ok(r.isError, `${tool} should reject ${p}`);
      assert.match(text(r), /Access denied/, `${tool} ${p}`);
    }
  }
  assert.match(text(await call("docx_info", { path: "report.pdf" })), /Unsupported file type/);
  assert.match(text(await call("docx_info", { path: "notes.txt" })), /Unsupported file type/);
  assert.match(text(await call("pdf_info", { path: "memo.docx" })), /Unsupported file type/);
  // A symlink that stays inside the allowed dir is fine.
  const ok = await call("pdf_info", { path: "sub/inside-link.pdf" });
  assert.ok(!ok.isError, text(ok));
  // Wrong types / missing files
  assert.match(text(await call("pdf_info", { path: "notes.txt" })), /Unsupported file type/);
  assert.match(text(await call("pdf_info", { path: "book.xlsx" })), /Unsupported file type/);
  assert.match(text(await call("pdf_info", { path: "missing.pdf" })), /not found/);
  assert.match(text(await call("pdf_info", { path: "sub" })), /Not a regular file|Unsupported/);
});

test("server ignores unexpanded placeholders and missing dirs", async () => {
  const stderr = getStderr();
  assert.match(stderr, /Allowed directories: .*allowed$/m);
  assert.match(stderr, /Cannot access allowed directory, ignoring: .*does-not-exist/);
});
});
