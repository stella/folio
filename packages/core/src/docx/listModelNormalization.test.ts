import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { normalizeListModelRendering } from "./listModelNormalization";
import { parseDocx } from "./parser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Paragraphs whose numbering the document defaults decide (a default
 * `w:numId` of 0 says "none") with no numbering part and no `w14:paraId`
 * anywhere, as some producers write.
 */
const styleNumberedWithoutIds = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId3" Type="${R}/header" Target="header1.xml"/></Relationships>`,
  );
  zip.set(
    "word/styles.xml",
    `${HEAD}<w:styles xmlns:w="${W}"><w:docDefaults><w:pPrDefault><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/></w:style></w:styles>`,
  );
  zip.set(
    "word/header1.xml",
    `${HEAD}<w:hdr xmlns:w="${W}"><w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>Header line</w:t></w:r></w:p></w:hdr>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14"><w:body><w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>Opening</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr><w:r><w:t>Second line</w:t></w:r></w:p><w:sectPr xmlns:r="${R}"><w:headerReference w:type="default" r:id="rId3"/></w:sectPr></w:body></w:document>`,
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer as ArrayBuffer;
};

describe("list normalization without paragraph identities", () => {
  test("a headless save writes the document", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await styleNumberedWithoutIds(),
    );
    const saved = await reviewer.toBuffer();
    const xml =
      (await (
        await ZipArchive.read(saved)
      )
        .entry("word/document.xml")
        ?.text()) ?? "";
    expect(xml).toContain("Second line");
    const header =
      (await (
        await ZipArchive.read(saved)
      )
        .entry("word/header1.xml")
        ?.text()) ?? "";
    expect(header).toContain("Header line");
  });

  test("normalizes without undo operations when none are asked for", async () => {
    const document = await parseDocx(await styleNumberedWithoutIds(), {
      preloadFonts: false,
    });
    for (const block of document.package.document.content)
      if (block.type === "paragraph") delete block.paraId;
    const { inverse } = normalizeListModelRendering(document, {
      inverse: "omit",
    });
    expect(inverse).toEqual([]);
  });
});
