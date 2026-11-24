import { expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { FolioDocxReviewer } from "./headless";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const controlDocx = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  );
  zip.set(
    "_rels/.rels",
    `${DECLARATION}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${DECLARATION}<w:document xmlns:w="${W}"><w:body>` +
      '<w:p><w:r><w:t xml:space="preserve">Pay </w:t></w:r>' +
      '<w:sdt><w:sdtPr><w:alias w:val="Amount"/><w:id w:val="7"/></w:sdtPr><w:sdtContent><w:r><w:t>100</w:t></w:r></w:sdtContent></w:sdt>' +
      '<w:r><w:t xml:space="preserve"> euros.</w:t></w:r></w:p>' +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

// A direct replacement of exactly a content control's text writes the new
// text in the control, which keeps its properties, as typing over it does.
for (const session of ["canonical"] as const)
  test(`${session}: replacing exactly an inline control's text keeps the control`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await controlDocx(), {
      session,
    });
    const [paragraph] = reviewer.snapshot().blocks;
    const result = reviewer.applyOperations(
      [
        {
          id: "op",
          type: "replaceInBlock",
          blockId: paragraph?.id ?? "",
          find: "100",
          replace: "250",
        },
      ],
      { undefinedReferences: "refuse", mode: "direct" },
    );
    expect(result.skipped).toEqual([]);
    const xml =
      (await (
        await ZipArchive.read(await reviewer.toBuffer())
      )
        .entry("word/document.xml")
        ?.text()) ?? "";
    const control = /<w:sdt>(.*?)<\/w:sdt>/u.exec(xml)?.[1] ?? "";
    expect(control).toContain('w:val="Amount"');
    expect(control).toContain("<w:t>250</w:t>");
    expect(xml).toContain("Pay ");
    expect(xml).toContain(" euros.");
  });
