/**
 * Field results in the AI-facing projection: a cross-reference must read as the
 * text Word shows, not vanish from the sentence around it.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { FolioDocxReviewer } from "./headless";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** `{ REF _Ref1 \r \h }` with a cached result, the shape Word writes. */
const referenceField = (result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> REF _Ref1 \\r \\h </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  (result === "" ? "" : `<w:r><w:t>${result}</w:t></w:r>`) +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const createFieldDocx = async (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
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
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">see Clause </w:t></w:r>${referenceField("3.6(a)")}<w:r><w:t xml:space="preserve"> above</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">see Clause </w:t></w:r>${referenceField("")}<w:r><w:t xml:space="preserve"> above</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">dated </w:t></w:r><w:fldSimple w:instr=" DOCPROPERTY Signed "><w:r><w:t>1 March</w:t></w:r></w:fldSimple></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`,
  );
  zip.file(
    "word/styles.xml",
    `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const blockTexts = async (): Promise<string[]> =>
  (await FolioDocxReviewer.fromBuffer(await createFieldDocx()))
    .getContent()
    .map(({ text }) => text);

describe("field results in AI block text", () => {
  test("reads a complex field as its result", async () => {
    expect((await blockTexts()).at(0)).toBe("see Clause 3.6(a) above");
  });

  test("adds nothing for a field with no result", async () => {
    expect((await blockTexts()).at(1)).toBe("see Clause  above");
  });

  test("reads a simple field as its content", async () => {
    expect((await blockTexts()).at(2)).toBe("dated 1 March");
  });

  test("keeps the annotated view on the same text", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await createFieldDocx());

    expect(reviewer.getContentAsText({ annotated: true })).toContain("see Clause 3.6(a) above");
  });
});
