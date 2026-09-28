// xlsx_edit: surgical OOXML editing. Proves that unrelated package parts survive
// byte-for-byte and that edits are correct, transactional and serialized.
// SERVER_ENTRY env var can point at an unpacked bundle's server/index.js.
import { test, before, after } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { connectLikeClaude } from "./claude-like-client.mjs";
import { buildRichWorkbook, buildNumbersOnly, buildImplicitRefs } from "./make-ooxml-fixture.mjs";

const here = path.dirname(fileURLToPath(import.meta.url));
const entry = path.resolve(process.env.SERVER_ENTRY || path.join(here, "..", "server", "index.js"));
const serverDir = path.dirname(entry);
const { readZip, readEntry } = await import(path.join(serverDir, "ooxml", "zip.js"));
const { applyEdits } = await import(path.join(serverDir, "xlsx-edit.js"));

const root = path.join(here, "fixtures", "ooxml");
const allowed = path.join(root, "allowed");
const outside = path.join(root, "outside");

const sha = (p) => crypto.createHash("sha256").update(fs.readFileSync(p)).digest("hex");
const text = (r) => r.content.filter((c) => c.type === "text").map((c) => c.text).join("\n");
const parts = (file) => {
  const buf = Buffer.isBuffer(file) ? file : fs.readFileSync(file);
  const z = readZip(buf);
  return new Map(z.entries.map((e) => [e.name, { raw: e.raw, entry: e }]));
};
const partText = (file, name) => readEntry(parts(file).get(name).entry).toString("utf8");
const load = async (file) => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  return wb;
};
const noLitter = (dir) => assert.deepEqual(fs.readdirSync(dir).filter((f) => f.startsWith(".~")), [], `temp/lock files left in ${dir}`);

/** Every entry of `before` not listed in `changed` must be raw-identical in `after` (and decompress identically). */
async function assertUntouched(beforeFile, afterFile, changed) {
  const b = parts(beforeFile);
  const a = parts(afterFile);
  const jb = await JSZip.loadAsync(fs.readFileSync(beforeFile));
  const ja = await JSZip.loadAsync(fs.readFileSync(afterFile), { checkCRC32: true });
  const actuallyChanged = [];
  for (const [name, e] of b) {
    assert.ok(a.has(name), `part ${name} missing after edit`);
    if (!e.raw.equals(a.get(name).raw)) actuallyChanged.push(name);
    else if (!jb.files[name].dir) {
      // Independent check with a different ZIP implementation.
      assert.ok(Buffer.from(await jb.files[name].async("uint8array")).equals(Buffer.from(await ja.files[name].async("uint8array"))), name);
    }
  }
  assert.deepEqual(actuallyChanged.sort(), [...changed].sort(), "exactly the expected parts changed");
  return { before: b, after: a };
}

let client;
const call = (name, args) => client.callTool({ name, arguments: args });
const edit = (args) => call("xlsx_edit", args);

before(async () => {
  fs.rmSync(root, { recursive: true, force: true });
  fs.mkdirSync(allowed, { recursive: true });
  fs.mkdirSync(outside, { recursive: true });
  fs.writeFileSync(path.join(allowed, "rich.xlsm"), await buildRichWorkbook({ macro: true }));
  fs.writeFileSync(path.join(allowed, "rich.xlsx"), await buildRichWorkbook({ macro: false }));
  fs.writeFileSync(path.join(allowed, "numbers.xlsx"), await buildNumbersOnly());
  fs.writeFileSync(path.join(allowed, "implicit.xlsx"), await buildImplicitRefs());
  fs.copyFileSync(path.join(allowed, "rich.xlsx"), path.join(outside, "secret.xlsx"));
  ({ client } = await connectLikeClaude(process.execPath, [entry, allowed], { env: { LOCAL_DOCUMENTS_LOCK_TIMEOUT_MS: "700" } }));
});
after(async () => client?.close());

test("xlsx_edit is the only workbook-editing tool and is not read-only", async () => {
  const { tools } = await client.listTools();
  assert.equal(tools.length, 16);
  assert.ok(!tools.some((t) => ["xlsx_write_range", "xlsx_set_formula", "xlsx_add_sheet"].includes(t.name)));
  const writers = ["xlsx_edit", "create_folder", "rename", "move", "move_batch"];
  for (const t of tools) assert.equal(t.annotations.readOnlyHint, !writers.includes(t.name), t.name);
});

test("a small edit leaves every unrelated part byte-identical (xlsm with add-ins, customXml, chart, pivot, VBA)", async () => {
  const src = path.join(allowed, "rich.xlsm");
  const srcHash = sha(src);
  const r = await edit({ path: "rich.xlsm", edits: [{ sheet: "Sheet1", cell: "B3", value: 177.5 }] });
  assert.ok(!r.isError, text(r));
  const out = path.join(allowed, "rich (edited).xlsm");
  assert.match(text(r), /Saved to new file: .*rich \(edited\)\.xlsm/);
  assert.match(text(r), /copied byte-for-byte/);
  assert.equal(sha(src), srcHash, "source unchanged");

  const { before: b, after: a } = await assertUntouched(src, out, ["xl/worksheets/sheet1.xml", "xl/workbook.xml"]);
  for (const name of [
    "xl/styles.xml", // custom table style
    "xl/webextensions/webextension1.xml",
    "xl/webextensions/taskpanes.xml",
    "xl/webextensions/_rels/taskpanes.xml.rels",
    "customXml/item1.xml",
    "customXml/itemProps1.xml",
    "customXml/_rels/item1.xml.rels",
    "docProps/custom.xml",
    "xl/charts/chart1.xml",
    "xl/drawings/drawing1.xml",
    "xl/pivotCache/pivotCacheDefinition1.xml",
    "xl/pivotCache/pivotCacheRecords1.xml",
    "xl/pivotTables/pivotTable1.xml",
    "xl/vbaProject.bin",
    "xl/calcChain.xml",
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/_rels/workbook.xml.rels",
    "xl/sharedStrings.xml",
  ]) {
    assert.ok(b.get(name).raw.equals(a.get(name).raw), `${name} must be byte-identical`);
  }
  assert.match(partText(out, "xl/styles.xml"), /tableStyle name="Corporate Blue"/);
  assert.match(partText(out, "docProps/custom.xml"), /ContentTypeId/);
  assert.equal(a.get("xl/vbaProject.bin").entry.method, 0, "VBA part still stored, not recompressed");

  // Inside the edited sheet, only the B3 cell element changed.
  const oldXml = partText(src, "xl/worksheets/sheet1.xml");
  const newXml = partText(out, "xl/worksheets/sheet1.xml");
  const oldCell = /<c r="B3"[^>]*>[\s\S]*?<\/c>/.exec(oldXml)[0];
  const style = /s="(\d+)"/.exec(oldCell)[1];
  assert.equal(newXml, oldXml.replace(oldCell, `<c r="B3" s="${style}"><v>177.5</v></c>`));
  // workbook.xml: only calcPr gained fullCalcOnLoad.
  const oldWb = partText(src, "xl/workbook.xml");
  const newWb = partText(out, "xl/workbook.xml");
  assert.equal(newWb.replace(/ fullCalcOnLoad="1"/, ""), oldWb);
  assert.match(newWb, /<calcPr[^>]*fullCalcOnLoad="1"/);

  const wb = await load(out);
  assert.equal(wb.getWorksheet("Sheet1").getCell("B3").value, 177.5);
  assert.equal(wb.getWorksheet("Sheet1").getCell("B3").numFmt, "#,##0.00", "style kept");
  noLitter(allowed);
});

test("value types, shared strings and inline strings", async () => {
  const r = await edit({
    path: "rich.xlsx",
    output_path: "types.xlsx",
    edits: [
      { sheet: "Sheet1", cell: "A2", value: "South" }, // already in sharedStrings -> reused
      { sheet: "Sheet1", cell: "A20", value: "Brand new & <special>" },
      { sheet: "Sheet1", cell: "A21", value: "  padded  " },
      { sheet: "Sheet1", cell: "A22", value: "_x0041_ literal" },
      { sheet: "Sheet1", cell: "B20", value: -0 },
      { sheet: "Sheet1", cell: "B21", value: 1e21 },
      { sheet: "Sheet1", cell: "B22", value: false },
      { sheet: "Sheet1", cell: "E2", value: { date: "2026-02-03" } },
      { sheet: "Sheet1", cell: "C20", value: "=not a formula" },
    ],
  });
  assert.ok(!r.isError, text(r));
  const out = path.join(allowed, "types.xlsx");
  const sheet = partText(out, "xl/worksheets/sheet1.xml");
  const sst = partText(out, "xl/sharedStrings.xml");
  const sstOld = partText(path.join(allowed, "rich.xlsx"), "xl/sharedStrings.xml");
  const southIndex = [...sstOld.matchAll(/<si>([\s\S]*?)<\/si>/g)].findIndex((m) => />South</.test(m[1]));
  assert.match(sheet, new RegExp(`<c r="A2" t="s"><v>${southIndex}</v></c>`));
  const uc = (x) => +/uniqueCount="(\d+)"/.exec(x)[1];
  assert.equal(uc(sst), uc(sstOld) + 4, "4 new unique strings appended");
  assert.match(sst, /<si><t>Brand new &amp; &lt;special&gt;<\/t><\/si>/);
  assert.match(sst, /<si><t xml:space="preserve">  padded  <\/t><\/si>/);
  assert.match(sst, /_x005F_x0041_ literal/);
  assert.match(sheet, /<c r="B20"><v>0<\/v><\/c>/);
  assert.match(sheet, /<c r="B22" t="b"><v>0<\/v><\/c>/);
  assert.match(sheet, /<c r="E2" s="\d+"><v>46056<\/v><\/c>/);
  const s = (await load(out)).getWorksheet("Sheet1");
  assert.equal(s.getCell("A20").value, "Brand new & <special>");
  assert.equal(s.getCell("A21").value, "  padded  ");
  assert.equal(s.getCell("B21").value, 1e21);
  assert.equal(s.getCell("B22").value, false);
  assert.equal(s.getCell("E2").value.toISOString().slice(0, 10), "2026-02-03");
  assert.equal(s.getCell("C20").value, "=not a formula");

  // Reusing an existing string only: sharedStrings.xml untouched.
  const r2 = await edit({ path: "rich.xlsx", output_path: "reuse.xlsx", edits: [{ sheet: "Sheet1", cell: "A3", value: "West" }] });
  assert.ok(!r2.isError, text(r2));
  await assertUntouched(path.join(allowed, "rich.xlsx"), path.join(allowed, "reuse.xlsx"), ["xl/worksheets/sheet1.xml", "xl/workbook.xml"]);

  // No sharedStrings part: inline string, and no new part is created.
  const r3 = await edit({ path: "numbers.xlsx", output_path: "numbers-out.xlsx", edits: [{ sheet: "Data", cell: "B1", value: "inline & text" }] });
  assert.ok(!r3.isError, text(r3));
  const n = path.join(allowed, "numbers-out.xlsx");
  assert.deepEqual([...parts(n).keys()].sort(), [...parts(path.join(allowed, "numbers.xlsx")).keys()].sort());
  assert.match(partText(n, "xl/worksheets/sheet1.xml"), /<c r="B1" t="inlineStr"><is><t>inline &amp; text<\/t><\/is><\/c>/);
  assert.equal((await load(n)).getWorksheet("Data").getCell("B1").value, "inline & text");
});

test("formulas: stored without cached value, escaped, and flagged for recalculation", async () => {
  const r = await edit({
    path: "rich.xlsx",
    output_path: "formulas.xlsx",
    edits: [
      { sheet: "Summary", cell: "B2", formula: "=COUNTA(Sheet1!A:A)" },
      { sheet: "Summary", cell: "B3", formula: 'IF(Sheet1!B2<5,"a&b","")' },
      { sheet: "R&D Plan", cell: "A2", formula: "A1*2" },
      { sheet: "Sheet1", cell: "D4", formula: "B4*C4" }, // replaces a shared-formula dependent: still a formula
    ],
  });
  assert.ok(!r.isError, text(r));
  const out = path.join(allowed, "formulas.xlsx");
  assert.match(partText(out, "xl/worksheets/sheet2.xml"), /<c r="B2"><f>COUNTA\(Sheet1!A:A\)<\/f><\/c>/);
  assert.match(partText(out, "xl/worksheets/sheet2.xml"), /<f>IF\(Sheet1!B2&lt;5,"a&amp;b",""\)<\/f>/);
  assert.match(partText(out, "xl/worksheets/sheet1.xml"), /<c r="D4"><f>B4\*C4<\/f><\/c>/);
  assert.match(partText(out, "xl/workbook.xml"), /fullCalcOnLoad="1"/);
  assert.ok(parts(out).get("xl/calcChain.xml").raw.equals(parts(path.join(allowed, "rich.xlsx")).get("xl/calcChain.xml").raw), "calcChain untouched when formulas stay formulas");
  const wb = await load(out);
  assert.equal(wb.getWorksheet("R&D Plan").getCell("A2").formula, "A1*2");
  assert.equal(wb.getWorksheet("Sheet1").getCell("D5").formula, "B5+C5", "other shared-formula cells intact");
});

test("clear keeps formatting; new cells inherit row/column styles", async () => {
  const r = await edit({
    path: "rich.xlsx",
    output_path: "clear.xlsx",
    edits: [
      { sheet: "Sheet1", cell: "B4", clear: true }, // styled
      { sheet: "Sheet1", cell: "C2", value: null }, // unstyled -> element removed
      { sheet: "Sheet1", cell: "F30", value: 5 }, // column F has a style
      { sheet: "Sheet1", cell: "C8", value: 6 }, // row 8 has a style
    ],
  });
  assert.ok(!r.isError, text(r));
  const xml = partText(path.join(allowed, "clear.xlsx"), "xl/worksheets/sheet1.xml");
  assert.match(xml, /<c r="B4" s="\d+"\/>/);
  assert.doesNotMatch(xml, /<c r="C2"/);
  const colStyle = /<col [^>]*min="6"[^>]*style="(\d+)"/.exec(xml)?.[1] ?? /<col [^>]*style="(\d+)"[^>]*min="6"/.exec(xml)?.[1];
  assert.ok(colStyle, "fixture has a column style");
  assert.match(xml, new RegExp(`<c r="F30" s="${colStyle}"><v>5</v></c>`));
  const rowS = /<row r="8"[^>]*\ss="(\d+)"/.exec(xml)[1];
  assert.match(xml, new RegExp(`<c r="C8" s="${rowS}"><v>6</v></c>`));
  const s = (await load(path.join(allowed, "clear.xlsx"))).getWorksheet("Sheet1");
  assert.equal(s.getCell("B4").value, null);
  assert.equal(s.getCell("B4").numFmt, "#,##0.00");
  assert.equal(s.getCell("F30").font?.italic, true);
  assert.equal(s.getCell("C8").font?.bold, true);
});

test("overwriting a shared-formula master promotes a dependent and prunes calcChain", async () => {
  const r = await edit({ path: "rich.xlsx", output_path: "master.xlsx", edits: [{ sheet: "Sheet1", cell: "D2", value: 0 }] });
  assert.ok(!r.isError, text(r));
  assert.match(text(r), /D3 became the master of shared formula/);
  assert.match(text(r), /Removed 1 stale calculation-chain entry/);
  const out = path.join(allowed, "master.xlsx");
  assert.match(partText(out, "xl/worksheets/sheet1.xml"), /<c r="D3"><f t="shared" si="0" ref="D3:D6">B3\+C3<\/f>/);
  const chain = partText(out, "xl/calcChain.xml");
  assert.doesNotMatch(chain, /r="D2"/);
  assert.match(chain, /<c r="D3" i="1"\/>/, "sheet id made explicit on the new first entry");
  const s = (await load(out)).getWorksheet("Sheet1");
  assert.equal(s.getCell("D2").value, 0);
  assert.equal(s.getCell("D6").formula, "B6+C6");
});

test("calcChain part is removed (with its relationship and content type) when it would become empty", async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet("S").getCell("A1").value = { formula: "1+1", result: 2 };
  const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
  zip.file("xl/calcChain.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>');
  let rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  zip.file("xl/_rels/workbook.xml.rels", rels.replace("</Relationships>", '<Relationship Id="rIdCc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>'));
  let ct = await zip.file("[Content_Types].xml").async("string");
  zip.file("[Content_Types].xml", ct.replace("</Types>", '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>'));
  const res = applyEdits(await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }), [{ sheet: "S", cell: "A1", value: 3 }]);
  const out = parts(res.buffer);
  assert.ok(!out.has("xl/calcChain.xml"));
  assert.doesNotMatch(readEntry(out.get("xl/_rels/workbook.xml.rels").entry).toString(), /calcChain/);
  assert.doesNotMatch(readEntry(out.get("[Content_Types].xml").entry).toString(), /calcChain/);
});

test("add_sheet updates workbook.xml, its rels and content types only; the new sheet is editable in the same batch", async () => {
  const r = await edit({
    path: "rich.xlsm",
    output_path: "added.xlsm",
    edits: [{ add_sheet: "Q3 Report" }, { sheet: "Q3 Report", cell: "A1", value: 42 }, { sheet: "q3 report", cell: "B1", formula: "A1*2" }],
  });
  assert.ok(!r.isError, text(r));
  const src = path.join(allowed, "rich.xlsm");
  const out = path.join(allowed, "added.xlsm");
  const { after } = await assertUntouched(src, out, ["xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"]);
  assert.ok(after.has("xl/worksheets/sheet5.xml"));
  assert.match(partText(out, "xl/workbook.xml"), /<sheet name="Q3 Report" sheetId="5" r:id="rId\d+"\/><\/sheets>/);
  assert.match(partText(out, "xl/_rels/workbook.xml.rels"), /Target="worksheets\/sheet5\.xml"/);
  assert.match(partText(out, "[Content_Types].xml"), /<Override PartName="\/xl\/worksheets\/sheet5\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.worksheet\+xml"\/>/);
  const wb = await load(out);
  assert.deepEqual(wb.worksheets.map((w) => w.name), ["Sheet1", "Summary", "R&D Plan", "Pivot", "Q3 Report"]);
  assert.equal(wb.getWorksheet("Q3 Report").getCell("A1").value, 42);
  assert.equal(wb.getWorksheet("Q3 Report").getCell("B1").formula, "A1*2");

  for (const [name, re] of [["sheet1", /already exists/], ["a/b", /cannot contain/], ["x".repeat(32), /1–31/], ["History", /reserved/], ["'q", /apostrophe/]]) {
    const bad = await edit({ path: "rich.xlsx", output_path: "never.xlsx", edits: [{ add_sheet: name }] });
    assert.ok(bad.isError, name);
    assert.match(text(bad), re);
  }
});

test("guards: merged cells and array formulas are refused", async () => {
  for (const [cell, re] of [["H1", /inside merged range G1:H1/], ["M2", /part of an array formula \(M1:M2\)/]]) {
    const r = await edit({ path: "rich.xlsx", output_path: "never.xlsx", edits: [{ sheet: "Sheet1", cell, value: 1 }] });
    assert.ok(r.isError, cell);
    assert.match(text(r), re);
  }
  const ok = await edit({ path: "rich.xlsx", output_path: "table-body.xlsx", edits: [{ sheet: "Sheet1", cell: "K2", value: 99 }, { sheet: "Sheet1", cell: "G1", value: "anchor ok" }] });
  assert.ok(!ok.isError, text(ok));
});

test("editing a table header cell renames the table column in both places", async () => {
  const src = path.join(allowed, "rich.xlsx");
  const r = await edit({ path: "rich.xlsx", output_path: "renamed-col.xlsx", edits: [{ sheet: "Sheet1", cell: "K1", value: "Value & <Notes>" }] });
  assert.ok(!r.isError, text(r));
  assert.match(text(r), /Renamed column "Val" of table "T1" to "Value & <Notes>"/);
  const out = path.join(allowed, "renamed-col.xlsx");
  const tablePart = [...parts(out).keys()].find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
  await assertUntouched(src, out, ["xl/worksheets/sheet1.xml", "xl/workbook.xml", "xl/sharedStrings.xml", tablePart].filter((n) => parts(src).has(n)));
  const tx = partText(out, tablePart);
  assert.match(tx, /<tableColumn [^>]*name="Key"/);
  assert.match(tx, /<tableColumn [^>]*name="Value &amp; &lt;Notes&gt;"/);
  const ws = (await load(out)).getWorksheet("Sheet1");
  assert.equal(ws.getCell("K1").value, "Value & <Notes>");
  assert.equal(ws.getCell("K2").value, 1, "table data unchanged");
});

test("table header renames are refused when unsafe", async () => {
  const cases = [
    [[{ sheet: "Sheet1", cell: "K1", value: 5 }], /must be text/],
    [[{ sheet: "Sheet1", cell: "K1", clear: true }], /must be text/],
    [[{ sheet: "Sheet1", cell: "K1", value: "  " }], /cannot be blank/],
    [[{ sheet: "Sheet1", cell: "K1", value: "key" }], /two columns named "key"/],
    [[{ sheet: "Sheet1", cell: "K1", value: "Amount" }, { sheet: "Sheet1", cell: "A20", formula: "SUM(T1[Val])" }], /Column "Val" of table "T1" is used by name in a formula/],
  ];
  for (const [edits, re] of cases) {
    const r = await edit({ path: "rich.xlsx", output_path: "never.xlsx", edits });
    assert.ok(r.isError, JSON.stringify(edits));
    assert.match(text(r), re);
  }
  // A formula already in the workbook blocks the rename too.
  const w = await edit({ path: "rich.xlsx", output_path: "with-ref.xlsx", edits: [{ sheet: "Summary", cell: "B1", formula: "SUM(T1[[#Data],[Val]])" }] });
  assert.ok(!w.isError, text(w));
  const r = await edit({ path: "with-ref.xlsx", output_path: "never.xlsx", edits: [{ sheet: "Sheet1", cell: "K1", value: "Amount" }] });
  assert.match(text(r), /is used by name in a formula/);
  // Swapping two names in one batch is fine.
  const ok = await edit({ path: "rich.xlsx", output_path: "swapped.xlsx", edits: [{ sheet: "Sheet1", cell: "J1", value: "Val" }, { sheet: "Sheet1", cell: "K1", value: "Key" }] });
  assert.ok(!ok.isError, text(ok));
  assert.ok(!fs.existsSync(path.join(allowed, "never.xlsx")));
});

test("transactional: one invalid edit means nothing is written", async () => {
  const src = path.join(allowed, "rich.xlsx");
  const h = sha(src);
  const cases = [
    [[{ sheet: "Sheet1", cell: "A1", value: 1 }, { sheet: "Nope", cell: "A1", value: 1 }], /Sheet "Nope" not found/],
    [[{ sheet: "Sheet1", cell: "A1", value: 1 }, { sheet: "Sheet1", cell: "H1", value: 1 }], /merged range/],
    [[{ sheet: "Sheet1", cell: "A1", value: 1, formula: "1" }], /exactly one of/],
    [[{ sheet: "Sheet1", cell: "ZZZZ1", value: 1 }], /Invalid cell reference|outside/],
    [[{ sheet: "Sheet1", cell: "A1", value: { date: "2026-02-30" } }], /Invalid date/],
    [[{ sheet: "Sheet1", cell: "A1", formula: "SUM(A1" }], /unbalanced parentheses/],
    [[{ add_sheet: "X", sheet: "Sheet1" }], /add_sheet cannot be combined|Invalid|Unrecognized|unrecognized/],
    [[{ sheet: "Sheet1", cell: "A1", valu: 1 }], /exactly one of|Invalid|Unrecognized|unrecognized/],
  ];
  for (const [edits, re] of cases) {
    const r = await edit({ path: "rich.xlsx", output_path: "never.xlsx", edits });
    assert.ok(r.isError, JSON.stringify(edits));
    assert.match(text(r), re);
  }
  assert.ok(!fs.existsSync(path.join(allowed, "never.xlsx")));
  assert.equal(sha(src), h);
  noLitter(allowed);
});

test("overwrite rules and output-path security", async () => {
  const copy = path.join(allowed, "inplace.xlsm");
  fs.copyFileSync(path.join(allowed, "rich.xlsm"), copy);
  const orig = path.join(root, "inplace-orig.xlsm");
  fs.copyFileSync(copy, orig);
  let r = await edit({ path: "inplace.xlsm", edits: [{ sheet: "Sheet1", cell: "C3", value: 7 }], overwrite: true });
  assert.ok(!r.isError, text(r));
  assert.match(text(r), /Saved over the source file/);
  await assertUntouched(orig, copy, ["xl/worksheets/sheet1.xml", "xl/workbook.xml"]);

  r = await edit({ path: "rich.xlsx", output_path: "rich.xlsx", edits: [{ sheet: "Sheet1", cell: "C3", value: 7 }] });
  assert.ok(r.isError);
  assert.match(text(r), /already exists.*overwrite=true/);

  for (const [output_path, re] of [
    ["out.xlsx", /extensions: \.xlsm/], // xlsm source must stay xlsm
    [path.join(outside, "x.xlsm"), /outside the allowed directories/],
    ["../outside/x.xlsm", /outside the allowed directories/],
    ["missing/x.xlsm", /Output folder does not exist/],
  ]) {
    const res = await edit({ path: "rich.xlsm", output_path, overwrite: true, edits: [{ sheet: "Sheet1", cell: "A1", value: 1 }] });
    assert.ok(res.isError, output_path);
    assert.match(text(res), re);
  }
  const src = await edit({ path: path.join(outside, "secret.xlsx"), edits: [{ sheet: "Sheet1", cell: "A1", value: 1 }] });
  assert.match(text(src), /Access denied/);
  assert.deepEqual(fs.readdirSync(outside), ["secret.xlsx"]);
  noLitter(allowed);
});

test("concurrent edits to one file are serialized; none are lost", async () => {
  const file = path.join(allowed, "concurrent.xlsx");
  fs.copyFileSync(path.join(allowed, "rich.xlsx"), file);
  const results = await Promise.all(
    Array.from({ length: 8 }, (_, i) => edit({ path: "concurrent.xlsx", overwrite: true, edits: [{ sheet: "Summary", cell: `C${i + 1}`, value: i + 1 }] }))
  );
  for (const r of results) assert.ok(!r.isError, text(r));
  const s = (await load(file)).getWorksheet("Summary");
  for (let i = 0; i < 8; i++) assert.equal(s.getCell(`C${i + 1}`).value, i + 1, `edit ${i + 1} kept`);

  // Parallel default-output edits get distinct file names.
  const outs = await Promise.all(Array.from({ length: 4 }, () => edit({ path: "concurrent.xlsx", edits: [{ sheet: "Summary", cell: "D1", value: 1 }] })));
  const names = outs.map((r) => /Saved to new file: (.*)/.exec(text(r))[1]);
  assert.equal(new Set(names).size, 4, names.join(" | "));
  noLitter(allowed);
});

test("a lock held by another live process blocks the edit; a stale lock is recovered", async () => {
  const file = path.join(allowed, "locked.xlsx");
  fs.copyFileSync(path.join(allowed, "rich.xlsx"), file);
  const lock = path.join(allowed, ".~localdocs-lock.locked.xlsx");
  fs.writeFileSync(lock, JSON.stringify({ pid: process.pid, host: os.hostname(), time: Date.now() }));
  const r = await edit({ path: "locked.xlsx", overwrite: true, edits: [{ sheet: "Sheet1", cell: "A1", value: "x" }] });
  assert.ok(r.isError);
  assert.match(text(r), /being edited by another process/);
  fs.writeFileSync(lock, JSON.stringify({ pid: 2 ** 22 + 12345, host: os.hostname(), time: Date.now() })); // dead pid
  const r2 = await edit({ path: "locked.xlsx", overwrite: true, edits: [{ sheet: "Sheet1", cell: "A1", value: "x" }] });
  assert.ok(!r2.isError, text(r2));
  assert.ok(!fs.existsSync(lock));
});

test("rows and cells without explicit r attributes are handled", async () => {
  const r = await edit({ path: "implicit.xlsx", output_path: "implicit-out.xlsx", edits: [{ sheet: "Data", cell: "A4", value: 4 }, { sheet: "Data", cell: "B2", value: 9 }] });
  assert.ok(!r.isError, text(r));
  // ExcelJS cannot read implicit cell positions (neither source nor result), so check the XML.
  assert.match(text(r), /semantic read-back check was skipped/);
  const xml = partText(path.join(allowed, "implicit-out.xlsx"), "xl/worksheets/sheet1.xml");
  assert.match(
    xml,
    /<sheetData><row r="1"><c><v>1<\/v><\/c><c><v>2<\/v><\/c><\/row><row r="2"><c r="A2"><v>3<\/v><\/c><c r="B2"><v>9<\/v><\/c><\/row><row r="4"><c r="A4"><v>4<\/v><\/c><\/row><row r="5"><c><v>5<\/v><\/c><\/row><\/sheetData>/
  );
});
