// Generates test/fixtures/allowed/memo.docx (+ an outside copy) with headings, nested
// lists, numbered lists, tables with merged cells, a footnote, an image and properties.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PNG } from "pngjs";
import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  VerticalMergeType, ExternalHyperlink, ImageRun, FootnoteReferenceRun, LevelFormat, AlignmentType, Header,
  CommentRangeStart, CommentRangeEnd, CommentReference,
} from "docx";

const here = path.dirname(fileURLToPath(import.meta.url));
const allowed = path.join(here, "fixtures", "allowed");
const outside = path.join(here, "fixtures", "outside");

const png = new PNG({ width: 8, height: 8 });
png.data.fill(200);
const cell = (text, opts = {}) => new TableCell({ children: [new Paragraph(text)], ...opts });

const doc = new Document({
  title: "Project Memo",
  creator: "Jane Tester",
  description: "Fixture for docx tools",
  subject: "Testing",
  keywords: "memo, fixture",
  lastModifiedBy: "Reviewer",
  comments: { children: [{ id: 0, author: "Rita Reviewer", date: new Date(Date.UTC(2026, 8, 1)), children: [new Paragraph("Please double-check this figure.")] }] },
  footnotes: { 1: { children: [new Paragraph("Source: internal audit 2026.")] } },
  numbering: {
    config: [{
      reference: "steps",
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.START },
        { level: 1, format: LevelFormat.LOWER_LETTER, text: "%2)", alignment: AlignmentType.START },
      ],
    }],
  },
  sections: [{
    headers: { default: new Header({ children: [new Paragraph("CONFIDENTIAL HEADER")] }) },
    children: [
      new Paragraph({ text: "Project Memo", heading: HeadingLevel.TITLE }),
      new Paragraph({ text: "Introduction", heading: HeadingLevel.HEADING_1 }),
      new Paragraph({ children: [
        new TextRun("The budget was "), new TextRun({ text: "approved", bold: true }),
        new TextRun(" on Monday."), new FootnoteReferenceRun(1),
      ] }),
      new Paragraph({ children: [
        new TextRun("See "),
        new ExternalHyperlink({ link: "https://example.com/plan", children: [new TextRun({ text: "the plan", style: "Hyperlink" })] }),
        new TextRun(" for details."),
      ] }),
      new Paragraph({ text: "Goals", heading: HeadingLevel.HEADING_2 }),
      new Paragraph({ text: "Reduce costs", bullet: { level: 0 } }),
      new Paragraph({ text: "Cloud spend", bullet: { level: 1 } }),
      new Paragraph({ text: "Licences", bullet: { level: 1 } }),
      new Paragraph({ text: "Improve quality", bullet: { level: 0 } }),
      new Paragraph({ text: "Steps", heading: HeadingLevel.HEADING_2 }),
      new Paragraph({ text: "Collect data", numbering: { reference: "steps", level: 0 } }),
      new Paragraph({ text: "Interview teams", numbering: { reference: "steps", level: 1 } }),
      new Paragraph({ text: "Write report", numbering: { reference: "steps", level: 0 } }),
      new Paragraph({ text: "Figures", heading: HeadingLevel.HEADING_1 }),
      new Paragraph({ children: [new ImageRun({ type: "png", data: PNG.sync.write(png), transformation: { width: 8, height: 8 }, altText: { name: "chart", description: "Revenue chart", title: "Chart" } })] }),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({ tableHeader: true, children: [cell("Region"), cell("Q1"), cell("Q2")] }),
          new TableRow({ children: [cell("North", { verticalMerge: VerticalMergeType.RESTART }), cell("100"), cell("150")] }),
          new TableRow({ children: [cell("", { verticalMerge: VerticalMergeType.CONTINUE }), cell("110"), cell("160")] }),
          new TableRow({ children: [cell("Total (needle)", { columnSpan: 2 }), cell("520")] }),
        ],
      }),
      new Paragraph({ children: [new CommentRangeStart(0), new TextRun("Text between tables."), new CommentRangeEnd(0), new TextRun({ children: [new CommentReference(0)] })] }),
      new Table({ rows: [
        new TableRow({ children: [cell("Key"), cell("Value")] }),
        new TableRow({ children: [cell("Owner"), cell("Finance\tteam")] }),
      ] }),
      new Paragraph({ text: "Conclusion", heading: HeadingLevel.HEADING_1 }),
      new Paragraph("Ship it. Unicode: café – 東京."),
    ],
  }],
});
const buf = await Packer.toBuffer(doc);
fs.writeFileSync(path.join(allowed, "memo.docx"), buf);
fs.writeFileSync(path.join(outside, "secret.docx"), buf);
try { fs.symlinkSync(path.join(outside, "secret.docx"), path.join(allowed, "link-to-secret.docx")); } catch {}
fs.writeFileSync(path.join(allowed, "broken.docx"), "this is not a zip file");
console.log("docx fixtures written");
