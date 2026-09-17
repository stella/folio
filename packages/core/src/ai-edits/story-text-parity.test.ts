/**
 * One text projection for every story, loaded or not.
 *
 * A story's text used to come from two different walks: the model walk when the
 * story had not been opened for editing, and a ProseMirror walk once it had.
 * They disagreed about tabs, hard breaks, table cells, tracked deletions and
 * paragraph separators, so the same note read one way before it was loaded and
 * another after. These assert the two agree byte for byte on content that
 * exercises every one of those disagreements.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { FolioDocxReviewer } from "./headless";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const HEADER_RELATIONSHIP_ID = "rId20";
const FOOTER_RELATIONSHIP_ID = "rId21";
const RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const complexField = (instruction: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  (result === "" ? "" : `<w:r><w:t xml:space="preserve">${result}</w:t></w:r>`) +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

/**
 * Every construct the two walks used to read differently: a tab and a hard
 * break (the model walk emitted them, the editor walk did not), a field with a
 * stored result and one without, a run nested inside a hyperlink, and a tracked
 * deletion (the model walk hides it, the editor walk showed it).
 */
const divergentRuns = (label: string) =>
  `<w:r><w:t xml:space="preserve">${label}</w:t></w:r>` +
  `<w:r><w:tab/></w:r>` +
  `<w:r><w:t xml:space="preserve">after tab</w:t></w:r>` +
  `<w:r><w:br/></w:r>` +
  `<w:r><w:t xml:space="preserve">after break</w:t></w:r>` +
  `<w:hyperlink r:id="rId99"><w:r><w:t xml:space="preserve"> nested</w:t></w:r></w:hyperlink>` +
  `<w:r><w:t xml:space="preserve"> page </w:t></w:r>${complexField("PAGE", "7")}` +
  `<w:r><w:t xml:space="preserve"> of </w:t></w:r>${complexField("PAGE", "")}` +
  `<w:del w:id="900" w:author="A" w:date="2020-01-01T00:00:00Z">` +
  `<w:r><w:delText xml:space="preserve"> removed</w:delText></w:r></w:del>` +
  `<w:r><w:t xml:space="preserve"> end</w:t></w:r>`;

const divergentParagraphs = (label: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${label} first</w:t></w:r></w:p>` +
  `<w:p>${divergentRuns(label)}</w:p>`;

const createStoryDocx = async (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}
<Relationships xmlns="${RELATIONSHIP_NAMESPACE}">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}
<Relationships xmlns="${RELATIONSHIP_NAMESPACE}">
  <Relationship Id="${HEADER_RELATIONSHIP_ID}" Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"/>
  <Relationship Id="${FOOTER_RELATIONSHIP_ID}" Type="${RELATIONSHIP_TYPES.footer}" Target="footer1.xml"/>
  <Relationship Id="rId30" Type="${RELATIONSHIP_TYPES.footnotes}" Target="footnotes.xml"/>
  <Relationship Id="rId31" Type="${RELATIONSHIP_TYPES.endnotes}" Target="endnotes.xml"/>
  <Relationship Id="rId99" Type="${RELATIONSHIP_TYPES.hyperlink}" Target="https://example.invalid/" TargetMode="External"/>
</Relationships>`,
  );

  const namespaces =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    `xmlns:r="${RELATIONSHIP_NAMESPACE}" ` +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

  zip.file(
    "word/header1.xml",
    `${XML_DECLARATION}\n<w:hdr ${namespaces}>${divergentParagraphs("header")}</w:hdr>`,
  );
  zip.file(
    "word/footer1.xml",
    `${XML_DECLARATION}\n<w:ftr ${namespaces}>${divergentParagraphs("footer")}</w:ftr>`,
  );
  zip.file(
    "word/footnotes.xml",
    `${XML_DECLARATION}
<w:footnotes ${namespaces}>
  <w:footnote w:id="2"><w:p w14:paraId="42000002">${divergentRuns("footnote")}</w:p></w:footnote>
</w:footnotes>`,
  );
  zip.file(
    "word/endnotes.xml",
    `${XML_DECLARATION}
<w:endnotes ${namespaces}>
  <w:endnote w:id="2"><w:p w14:paraId="43000002">${divergentRuns("endnote")}</w:p></w:endnote>
</w:endnotes>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document ${namespaces}>
  <w:body>
    <w:p w14:paraId="41000001">${divergentRuns("body")}</w:p>
    <w:p w14:paraId="41000002"><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:t>note anchor</w:t></w:r></w:p>
    <w:p w14:paraId="41000003"><w:r><w:endnoteReference w:id="2"/></w:r><w:r><w:t>end anchor</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="${HEADER_RELATIONSHIP_ID}"/>
      <w:footerReference w:type="default" r:id="${FOOTER_RELATIONSHIP_ID}"/>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );
  zip.file(
    "word/styles.xml",
    `${XML_DECLARATION}\n<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/** Secondary stories only: the body has no unloaded projection to differ from. */
const secondaryStoryTexts = async (load: boolean): Promise<Map<string, string>> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await createStoryDocx());
  const handles = reviewer
    .listStories()
    .map(({ handle }) => handle)
    .filter((handle) => handle.type !== "main");
  if (load) {
    for (const handle of handles) {
      reviewer.snapshotStory(handle);
    }
  }
  return new Map(
    reviewer
      .listStories()
      .filter(({ handle }) => handle.type !== "main")
      .map(({ handle, text }) => [handle.type, text]),
  );
};

describe("story text parity", () => {
  test("reads a loaded story exactly as an unloaded one", async () => {
    const unloaded = await secondaryStoryTexts(false);
    const loaded = await secondaryStoryTexts(true);

    expect([...loaded.keys()].toSorted()).toEqual(["endnote", "footer", "footnote", "header"]);
    for (const [kind, text] of unloaded) {
      expect(loaded.get(kind)).toBe(text);
    }
  });

  test("keeps the separators a reader needs and hides a tracked deletion", async () => {
    const unloaded = await secondaryStoryTexts(false);

    // A tab and a break separate words rather than gluing them, the stored
    // field result reads through, the result-less field adds nothing, and the
    // accepted view drops the deleted run.
    expect(unloaded.get("footnote")).toBe("footnote after tab after break nested page 7 of end");
  });
});
