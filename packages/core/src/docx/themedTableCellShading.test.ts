import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { expectTableCellAttrs } from "../prosemirror/attrs";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const createThemedTableFixture = async (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.styles}" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.theme}" Target="theme/theme1.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/styles.xml",
    `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:styleId="ThemedTable">
    <w:name w:val="Themed Table"/>
    <w:tcPr>
      <w:shd w:val="clear" w:fill="70AD47" w:themeFill="accent6"/>
    </w:tcPr>
  </w:style>
</w:styles>`,
  );
  zip.file(
    "word/theme/theme1.xml",
    `${XML_DECLARATION}
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Theme table cell fills</w:t></w:r></w:p>
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="ThemedTable"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:shd w:val="clear" w:fill="DAE3F3" w:themeFill="accent1" w:themeFillTint="33"/></w:tcPr>
          <w:p><w:r><w:t>Direct tinted theme fill</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:p><w:r><w:t>Table style theme fill</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );

  return zip.generateAsync({ type: "arraybuffer" });
};

const tableCellBackgrounds = (document: ReturnType<typeof toProseDoc>): string[] => {
  const table = document.child(1);
  const row = table.firstChild;
  if (!row) {
    throw new Error("Expected themed table row");
  }

  return [
    expectTableCellAttrs(row.child(0)).backgroundColor ?? "",
    expectTableCellAttrs(row.child(1)).backgroundColor ?? "",
  ];
};

describe("themed table cell shading", () => {
  test("renders direct and style fills while preserving theme references on save", async () => {
    const original = await parseDocx(await createThemedTableFixture());
    const pmDocument = toProseDoc(original);

    expect(tableCellBackgrounds(pmDocument)).toEqual(["DAE3F3", "70AD47"]);

    const rebuilt = fromProseDoc(pmDocument, original);
    const repacked = await repackDocx(rebuilt, { updateModifiedDate: false });
    const repackedZip = await JSZip.loadAsync(repacked);
    const documentXml = await repackedZip.file("word/document.xml")?.async("string");

    expect(documentXml).toContain('w:themeFill="accent1"');
    expect(documentXml).toContain('w:themeFillTint="33"');

    const reopened = await parseDocx(repacked);
    expect(tableCellBackgrounds(toProseDoc(reopened))).toEqual(["DAE3F3", "70AD47"]);
  }, 15_000);
});
