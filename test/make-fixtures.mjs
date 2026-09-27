// Generates test fixtures in test/fixtures (and an "outside" dir for escape tests).
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { PNG } from "pngjs";
import ExcelJS from "exceljs";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(here, "fixtures");
const allowed = path.join(root, "allowed");
const outside = path.join(root, "outside");
const sibling = path.join(root, "allowed-evil"); // prefix-sharing sibling
fs.rmSync(root, { recursive: true, force: true });
for (const d of [allowed, path.join(allowed, "sub"), outside, sibling]) fs.mkdirSync(d, { recursive: true });

// Text PDF, 3 pages
{
  const doc = await PDFDocument.create();
  doc.setTitle("Quarterly Report");
  doc.setAuthor("Test Author");
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let i = 1; i <= 3; i++) {
    const page = doc.addPage([612, 792]);
    page.drawText(`Page ${i} heading`, { x: 72, y: 700, size: 20, font });
    page.drawText(`Revenue for region ${i} was ${i * 1000} units.`, { x: 72, y: 660, size: 12, font });
    page.drawText("Second line of body text.", { x: 72, y: 644, size: 12, font });
  }
  fs.writeFileSync(path.join(allowed, "report.pdf"), await doc.save());
}

// "Scanned" PDF: one page that is only an image (no text layer)
{
  const w = 400, h = 200;
  const png = new PNG({ width: w, height: h });
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = (y * w + x) * 4;
    const inBox = x > 50 && x < 350 && y > 50 && y < 150;
    png.data[i] = inBox ? 200 : 255; png.data[i + 1] = inBox ? 30 : 255; png.data[i + 2] = inBox ? 30 : 255; png.data[i + 3] = 255;
  }
  const doc = await PDFDocument.create();
  const img = await doc.embedPng(PNG.sync.write(png));
  const page = doc.addPage([612, 792]);
  page.drawImage(img, { x: 106, y: 296, width: 400, height: 200 });
  page.drawRectangle({ x: 50, y: 50, width: 100, height: 40, color: rgb(0, 0.4, 1) });
  fs.writeFileSync(path.join(allowed, "sub", "scanned.pdf"), await doc.save());
  fs.writeFileSync(path.join(outside, "secret.pdf"), await doc.save());
  fs.writeFileSync(path.join(sibling, "evil.pdf"), await doc.save());
}

// Workbook
{
  const wb = new ExcelJS.Workbook();
  wb.creator = "Fixture";
  const s = wb.addWorksheet("Sales");
  s.addRow(["Region", "Q1", "Q2", "Total", "Updated"]);
  s.addRow(["North", 100, 150, { formula: "B2+C2", result: 250 }, new Date(Date.UTC(2026, 0, 15))]);
  s.addRow(["South", 80.5, 20, { formula: "B3+C3", result: 100.5 }, new Date(Date.UTC(2026, 1, 1))]);
  s.addRow(["Multi\nline", true, { richText: [{ text: "rich " }, { text: "text" }] }, { formula: "1/0", result: { error: "#DIV/0!" } }, null]);
  s.mergeCells("A6:C6");
  s.getCell("A6").value = "Merged banner";
  const h = wb.addWorksheet("Hidden Notes", { state: "hidden" });
  h.getCell("B3").value = "needle in hidden sheet";
  h.getCell("J40").value = { text: "Example link", hyperlink: "https://example.com" };
  wb.definedNames.add("Sales!$B$2:$C$3", "Quarters");
  await wb.xlsx.writeFile(path.join(allowed, "book.xlsx"));
  await wb.xlsx.writeFile(path.join(outside, "secret.xlsx"));
}

// Escape attempts: symlinks pointing outside (skip where symlinks aren't permitted)
try {
  fs.symlinkSync(path.join(outside, "secret.pdf"), path.join(allowed, "link-to-secret.pdf"));
  fs.symlinkSync(outside, path.join(allowed, "link-dir"), "junction");
  fs.symlinkSync(path.join(allowed, "report.pdf"), path.join(allowed, "sub", "inside-link.pdf"));
} catch (e) {
  console.error("symlink creation skipped:", e.code);
}
fs.writeFileSync(path.join(allowed, "notes.txt"), "not a pdf");
console.log("fixtures in", root);
await import("./make-docx-fixture.mjs");
