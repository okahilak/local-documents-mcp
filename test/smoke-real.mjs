// Smoke test against real-world PDFs on this machine plus a large generated workbook.
import path from "node:path";
import fs from "node:fs";
import os from "node:os";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";

const entry = process.env.SERVER_ENTRY || path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "server", "index.js");
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "ld-smoke-"));
const big = path.join(tmp, "big.xlsx");
{
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Data");
  ws.addRow(["id", "name", "amount", "date"]);
  for (let i = 1; i <= 50000; i++) ws.addRow([i, `Customer ${i}`, i * 1.25, new Date(Date.UTC(2025, 0, 1 + (i % 365)))]);
  await wb.xlsx.writeFile(big);
}
const pdfs = process.argv.slice(2);
const dirs = [tmp, ...new Set(pdfs.map((p) => path.dirname(p)))];
const client = new Client({ name: "smoke", version: "1" });
await client.connect(new StdioClientTransport({ command: process.execPath, args: [entry, ...dirs], stderr: "ignore" }));
const call = async (name, args) => {
  const t0 = performance.now();
  const r = await client.callTool({ name, arguments: args });
  const ms = Math.round(performance.now() - t0);
  const txt = r.content.filter((c) => c.type === "text").map((c) => c.text).join("\n");
  const img = r.content.find((c) => c.type === "image");
  console.log(`\n### ${name} ${JSON.stringify(args).replace(tmp, "<tmp>")}  [${ms} ms]${r.isError ? " ERROR" : ""}`);
  console.log(txt.slice(0, 500) + (txt.length > 500 ? `\n… (${txt.length} chars)` : ""));
  if (img) console.log(`image: ${img.mimeType}, ${Math.round((img.data.length * 3) / 4 / 1024)} KB`);
  return r;
};
for (const p of pdfs) {
  const info = JSON.parse((await call("pdf_info", { path: p })).content[0].text);
  await call("pdf_read_text", { path: p, pages: "1" });
  const r = await call("pdf_render_page", { path: p, page: Math.min(2, info.pages) });
  const img = r.content.find((c) => c.type === "image");
  if (img) fs.writeFileSync(path.join(tmp, path.basename(path.dirname(p)) + ".img"), Buffer.from(img.data, "base64"));
}
await call("xlsx_info", { path: big });
await call("xlsx_read_range", { path: big, sheet: "Data", range: "A49995:D50001" });
await call("xlsx_read_range", { path: big, sheet: "Data", range: "A:D" });
await call("xlsx_search", { path: big, query: "Customer 4999", max_results: 5 });
console.log("\nimages saved in", tmp);
await client.close();
