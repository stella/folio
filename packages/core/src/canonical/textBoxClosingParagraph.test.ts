/**
 * A text box is a container, and a container ends with a paragraph: one whose
 * content ends in a table reads the empty paragraph after it, saves unedited
 * as read, and writes the paragraph once the box is edited.
 */
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";

const XML_DECLARATION =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const CONTENT_TYPES = `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
const ROOT_RELS = `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
const DOCUMENT_RELS = `${XML_DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'mc:Ignorable="w14"',
].join(" ");

const paragraph = (id: string, text: string) =>
  `<w:p w14:paraId="${id}" w14:textId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const TABLE = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${paragraph("20000003", "Cell")}</w:tc></w:tr></w:tbl>`;
const BOX =
  "<w:r><w:drawing><wp:inline>" +
  '<wp:extent cx="914400" cy="457200"/><wp:docPr id="11" name="Text Box 11"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  `<wps:txbx><w:txbxContent>${paragraph("20000002", "Box opens")}${TABLE}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
  "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
const DOCUMENT = `${XML_DECLARATION}<w:document ${NAMESPACES}><w:body><w:p w14:paraId="20000001" w14:textId="20000001"><w:r><w:t>Host</w:t></w:r>${BOX}</w:p>${paragraph("20000004", "After")}<w:sectPr/></w:body></w:document>`;

const buildPackage = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set("[Content_Types].xml", CONTENT_TYPES);
  zip.set("_rels/.rels", ROOT_RELS);
  zip.set("word/_rels/document.xml.rels", DOCUMENT_RELS);
  zip.set("word/document.xml", DOCUMENT);
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

const documentPart = async (bytes: ArrayBuffer): Promise<string> =>
  (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
  "";

const boxContent = (xml: string): string =>
  /<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/u.exec(xml)?.[1] ?? "";

describe("a text box whose content ends in a table", () => {
  test("saves unedited as read", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildPackage());
    expect(await documentPart(await reviewer.toBuffer())).toBe(DOCUMENT);
  });

  test("ends with a paragraph once the box is edited", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildPackage());
    const result = reviewer.applyOperations(
      [
        {
          id: "edit",
          type: "replaceInBlock",
          blockId: "20000002",
          find: "opens",
          replace: "starts",
        },
      ],
      { undefinedReferences: "refuse", mode: "direct" },
    );
    expect(result.skipped).toEqual([]);
    const content = boxContent(await documentPart(await reviewer.toBuffer()));
    expect(content).toContain("Box starts");
    expect(content).toMatch(/<\/w:tbl><w:p\b[^>]*\/?>(?:<\/w:p>)?$/u);
  });
});
