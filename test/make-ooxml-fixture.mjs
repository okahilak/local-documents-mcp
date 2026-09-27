// Builds realistic OOXML test workbooks: an ExcelJS-generated base, then extra parts that
// ExcelJS itself would drop on save (custom table style, Office add-in, SharePoint
// customXml, custom document properties, chart, pivot table, VBA project, calcChain).
// These are injected with JSZip here purely to CREATE fixtures.
import ExcelJS from "exceljs";
import JSZip from "jszip";
import crypto from "node:crypto";

const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

async function base() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.getColumn(6).font = { italic: true }; // column F style for inheritance tests
  ws.addRow(["Region", "Q1", "Q2", "Total", "Updated"]);
  ws.getRow(1).font = { bold: true, color: { argb: "FF1F4E79" } };
  const regions = ["North", "South", "East", "West", "Central"];
  regions.forEach((r, i) => {
    const row = i + 2;
    ws.getCell(`A${row}`).value = r;
    ws.getCell(`B${row}`).value = (i + 1) * 100;
    ws.getCell(`C${row}`).value = (i + 1) * 10;
    ws.getCell(`B${row}`).numFmt = "#,##0.00";
    ws.getCell(`E${row}`).value = new Date(Date.UTC(2026, 0, 10 + i));
    ws.getCell(`E${row}`).numFmt = "yyyy-mm-dd";
  });
  ws.getCell("D2").value = { formula: "B2+C2", shareType: "shared", ref: "D2:D6", result: 110 };
  for (let r = 3; r <= 6; r++) ws.getCell(`D${r}`).value = { sharedFormula: "D2", result: (r - 1) * 110 };
  ws.mergeCells("G1:H1");
  ws.getCell("G1").value = "Merged header";
  ws.getRow(8).font = { bold: true }; // row style for inheritance tests
  ws.getCell("A8").value = "Styled row";
  ws.getCell("M1").value = { formula: "B2:B3*2", shareType: "array", ref: "M1:M2", result: 200 };
  ws.addTable({ name: "T1", ref: "J1", columns: [{ name: "Key" }, { name: "Val" }], rows: [["a", 1], ["b", 2]] });
  wb.addWorksheet("Summary").getCell("A1").value = "Summary";
  wb.addWorksheet("R&D Plan").getCell("A1").value = 1;
  wb.addWorksheet("Pivot").getCell("A1").value = "pivot here";
  return wb.xlsx.writeBuffer();
}

function addRel(xml, id, type, target) {
  return xml.replace("</Relationships>", `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`);
}
function addOverride(xml, part, ct) {
  return xml.replace("</Types>", `<Override PartName="${part}" ContentType="${ct}"/></Types>`);
}

/** @param {{macro?: boolean}} opts */
export async function buildRichWorkbook({ macro = false } = {}) {
  const zip = await JSZip.loadAsync(await base());
  const text = (p) => zip.file(p).async("string");
  let ct = await text("[Content_Types].xml");
  let rootRels = await text("_rels/.rels");
  let wbRels = await text("xl/_rels/workbook.xml.rels");
  let wbXml = await text("xl/workbook.xml");

  // Custom table style (styles.xml)
  let styles = await text("xl/styles.xml");
  const tableStyles =
    '<tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"><tableStyle name="Corporate Blue" pivot="0" count="1"><tableStyleElement type="wholeTable" dxfId="0"/></tableStyle></tableStyles>';
  styles = styles.replace(/<tableStyles[^>]*\/>|<tableStyles[\s\S]*?<\/tableStyles>/, "");
  styles = styles.replace(/<dxfs[^>]*\/>|<dxfs[\s\S]*?<\/dxfs>/, "");
  styles = styles.replace("</cellStyles>", '</cellStyles><dxfs count="1"><dxf><font><b/><color rgb="FF1F4E79"/></font></dxf></dxfs>' + tableStyles);
  zip.file("xl/styles.xml", styles);

  // Office add-in (web extension task pane)
  zip.file(
    "xl/webextensions/webextension1.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{6F5B7E1C-0000-4000-8000-000000000001}"><we:reference id="wa104379955" version="1.0.0.0" store="en-US" storeType="OMEX"/><we:alternateReferences/><we:properties><we:property name="Office.AutoShowTaskpaneWithDocument" value="true"/></we:properties><we:bindings/><we:snapshot xmlns:r="' + NS_R + '"/></we:webextension>'
  );
  zip.file(
    "xl/webextensions/taskpanes.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"><wetp:taskpane dockstate="right" visibility="1" width="350" row="4"><wetp:webextensionref xmlns:r="' + NS_R + '" r:id="rId1"/></wetp:taskpane></wetp:taskpanes>'
  );
  zip.file(
    "xl/webextensions/_rels/taskpanes.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" Target="webextension1.xml"/></Relationships>'
  );
  rootRels = addRel(rootRels, "rIdWe", "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes", "xl/webextensions/taskpanes.xml");
  ct = addOverride(ct, "/xl/webextensions/taskpanes.xml", "application/vnd.ms-office.webextensiontaskpanes+xml");
  ct = addOverride(ct, "/xl/webextensions/webextension1.xml", "application/vnd.ms-office.webextension+xml");

  // SharePoint customXml (content type schema + item props)
  zip.file(
    "customXml/item1.xml",
    '<?xml version="1.0" encoding="utf-8"?><ct:contentTypeSchema ct:_="" ma:_="" ma:contentTypeName="Document" ma:contentTypeID="0x0101004D2F0E1234" ma:contentTypeVersion="12" xmlns:ct="http://schemas.microsoft.com/office/2006/metadata/contentType" xmlns:ma="http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes"></ct:contentTypeSchema>'
  );
  zip.file(
    "customXml/itemProps1.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="{A1B2C3D4-0000-4000-8000-000000000002}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="http://schemas.microsoft.com/office/2006/metadata/contentType"/></ds:schemaRefs></ds:datastoreItem>'
  );
  zip.file(
    "customXml/_rels/item1.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>'
  );
  wbRels = addRel(wbRels, "rIdCx1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml", "../customXml/item1.xml");
  ct = addOverride(ct, "/customXml/itemProps1.xml", "application/vnd.openxmlformats-officedocument.customXmlProperties+xml");

  // Custom document properties (incl. SharePoint ContentTypeId)
  zip.file(
    "docProps/custom.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ContentTypeId"><vt:lpwstr>0x0101004D2F0E1234</vt:lpwstr></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Project Code"><vt:lpwstr>LD-2026</vt:lpwstr></property></Properties>'
  );
  rootRels = addRel(rootRels, "rIdCustom", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties", "docProps/custom.xml");
  ct = addOverride(ct, "/docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml");

  // Chart (+ drawing) anchored on Sheet1
  zip.file(
    "xl/charts/chart1.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:f>Sheet1!$B$2:$B$6</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart></c:plotArea></c:chart></c:chartSpace>'
  );
  zip.file(
    "xl/drawings/drawing1.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>18</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="' + NS_R + '" r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'
  );
  zip.file(
    "xl/drawings/_rels/drawing1.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>'
  );
  ct = addOverride(ct, "/xl/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
  ct = addOverride(ct, "/xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml");
  // Hook the drawing into Sheet1 (sheet rels + <drawing r:id>)
  const s1RelsPath = "xl/worksheets/_rels/sheet1.xml.rels";
  let s1Rels = zip.file(s1RelsPath) ? await text(s1RelsPath) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  s1Rels = addRel(s1Rels, "rIdDr1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing", "../drawings/drawing1.xml");
  zip.file(s1RelsPath, s1Rels);
  let s1 = await text("xl/worksheets/sheet1.xml");
  s1 = s1.replace(/(<tableParts)/, `<drawing r:id="rIdDr1"/>$1`);
  zip.file("xl/worksheets/sheet1.xml", s1);

  // Pivot cache + pivot table on the "Pivot" sheet (sheet4)
  zip.file(
    "xl/pivotCache/pivotCacheDefinition1.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="' + NS_R + '" r:id="rId1" refreshOnLoad="1" recordCount="0"><cacheSource type="worksheet"><worksheetSource ref="A1:C6" sheet="Sheet1"/></cacheSource><cacheFields count="1"><cacheField name="Region" numFmtId="0"><sharedItems/></cacheField></cacheFields></pivotCacheDefinition>'
  );
  zip.file(
    "xl/pivotCache/pivotCacheRecords1.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>'
  );
  zip.file(
    "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>'
  );
  zip.file(
    "xl/pivotTables/pivotTable1.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="7" dataCaption="Values"><location ref="A3:A4" firstHeaderRow="1" firstDataRow="1" firstDataCol="0"/><pivotFields count="1"><pivotField showAll="0"/></pivotFields></pivotTableDefinition>'
  );
  zip.file(
    "xl/pivotTables/_rels/pivotTable1.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>'
  );
  zip.file(
    "xl/worksheets/_rels/sheet4.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>'
  );
  wbRels = addRel(wbRels, "rIdPc1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition", "pivotCache/pivotCacheDefinition1.xml");
  wbXml = wbXml.replace(/(<calcPr[^>]*\/>)/, `$1<pivotCaches><pivotCache cacheId="7" r:id="rIdPc1"/></pivotCaches>`);
  ct = addOverride(ct, "/xl/pivotCache/pivotCacheDefinition1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
  ct = addOverride(ct, "/xl/pivotCache/pivotCacheRecords1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml");
  ct = addOverride(ct, "/xl/pivotTables/pivotTable1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");

  // calcChain for the shared formula cells on Sheet1 (sheetId 1)
  zip.file(
    "xl/calcChain.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="D2" i="1"/><c r="D3"/><c r="D4"/><c r="D5"/><c r="D6"/><c r="M1" a="1"/></calcChain>'
  );
  wbRels = addRel(wbRels, "rIdCc1", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain", "calcChain.xml");
  ct = addOverride(ct, "/xl/calcChain.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml");

  // VBA project (macro-enabled variant); stored uncompressed like Excel often does
  if (macro) {
    const vba = Buffer.concat([Buffer.from("d0cf11e0a1b11ae1", "hex"), crypto.randomBytes(4096)]);
    zip.file("xl/vbaProject.bin", vba, { compression: "STORE" });
    wbRels = addRel(wbRels, "rIdVba", "http://schemas.microsoft.com/office/2006/relationships/vbaProject", "vbaProject.bin");
    ct = ct.replace("<Default ", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ');
    ct = ct.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", "application/vnd.ms-excel.sheet.macroEnabled.main+xml");
  }

  zip.file("[Content_Types].xml", ct);
  zip.file("_rels/.rels", rootRels);
  zip.file("xl/_rels/workbook.xml.rels", wbRels);
  zip.file("xl/workbook.xml", wbXml);
  // streamFiles => every entry uses a data descriptor, exercising verbatim copies.
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", streamFiles: true });
}

/** A workbook with numbers only (no sharedStrings part). */
export async function buildNumbersOnly() {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet("Data").getCell("A1").value = 1;
  const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

/** A workbook whose rows and cells omit the optional r attributes. */
export async function buildImplicitRefs() {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet("Data").getCell("A1").value = 1;
  const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
  let s = await zip.file("xl/worksheets/sheet1.xml").async("string");
  s = s.replace(/<sheetData>[\s\S]*<\/sheetData>|<sheetData\/>/, "<sheetData><row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c></row><row r=\"5\"><c><v>5</v></c></row></sheetData>");
  zip.file("xl/worksheets/sheet1.xml", s);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}
