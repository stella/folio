/**
 * Field results in the notes projection: a PAGE or DATE field in a footnote,
 * header or footer must read as the result stored in the package, never as the
 * placeholder the editor paints for an unresolved field — one branch of which
 * is the current date, so the same document read twice could otherwise
 * disagree.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { FolioDocxReviewer } from "./headless";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** A complex field with the cached result Word writes between separate and end. */
const complexField = (instruction: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  (result === "" ? "" : `<w:r><w:t xml:space="preserve">${result}</w:t></w:r>`) +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const footnote = (id: number, content: string) =>
  `<w:footnote w:id="${id}"><w:p w14:paraId="4200000${id}">${content}</w:p></w:footnote>`;

const PAGE_RESULT = "7";
const DATE_INSTRUCTION = 'DATE \\@ "d MMMM yyyy"';
const DATE_RESULT = "1 January 2020";
const PLAIN_NOTE_TEXT = "Plain note text.";

const NOTE_WITH_RESULTS_ID = 2;
const NOTE_WITHOUT_RESULT_ID = 3;
const NOTE_WITHOUT_FIELD_ID = 4;
const NOTE_WITH_CONTROL_ATOMS_ID = 5;
const NOTE_WITH_STRUCTURED_FIELD_ID = 6;

const RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEADER_RELATIONSHIP_ID = "rId20";
const FOOTER_RELATIONSHIP_ID = "rId21";

/**
 * A header or footer line carrying both cases: fields with a stored result, and
 * the result-less PAGE and DATE that the editor would otherwise paint as
 * `{page}` and today's date.
 */
const storyFieldRuns = (label: string) =>
  `<w:r><w:t xml:space="preserve">${label} page </w:t></w:r>${complexField("PAGE", PAGE_RESULT)}` +
  `<w:r><w:t xml:space="preserve"> dated </w:t></w:r>${complexField(DATE_INSTRUCTION, DATE_RESULT)}` +
  `<w:r><w:t xml:space="preserve"> of </w:t></w:r>${complexField("PAGE", "")}` +
  `<w:r><w:t xml:space="preserve"> </w:t></w:r>${complexField(DATE_INSTRUCTION, "")}` +
  `<w:r><w:t xml:space="preserve"> end</w:t></w:r>`;

const createNoteFieldDocx = async (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
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
  <Relationship Id="rId10" Type="${RELATIONSHIP_TYPES.footnotes}" Target="footnotes.xml"/>
  <Relationship Id="${HEADER_RELATIONSHIP_ID}" Type="${RELATIONSHIP_TYPES.header}" Target="header1.xml"/>
  <Relationship Id="${FOOTER_RELATIONSHIP_ID}" Type="${RELATIONSHIP_TYPES.footer}" Target="footer1.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${RELATIONSHIP_NAMESPACE}">
  <w:body>
    <w:p><w:r><w:t>Body paragraph.</w:t></w:r><w:r><w:footnoteReference w:id="${NOTE_WITH_RESULTS_ID}"/></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="${HEADER_RELATIONSHIP_ID}"/>
      <w:footerReference w:type="default" r:id="${FOOTER_RELATIONSHIP_ID}"/>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>
  </w:body>
</w:document>`,
  );
  zip.file(
    "word/header1.xml",
    `${XML_DECLARATION}
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p w14:paraId="42000010">${storyFieldRuns("header")}</w:p>
</w:hdr>`,
  );
  zip.file(
    "word/footer1.xml",
    `${XML_DECLARATION}
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:p w14:paraId="42000011">${storyFieldRuns("footer")}</w:p>
</w:ftr>`,
  );
  zip.file(
    "word/footnotes.xml",
    `${XML_DECLARATION}
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  ${footnote(
    NOTE_WITH_RESULTS_ID,
    `<w:r><w:t xml:space="preserve">page </w:t></w:r>${complexField("PAGE", PAGE_RESULT)}` +
      `<w:r><w:t xml:space="preserve"> dated </w:t></w:r>${complexField(DATE_INSTRUCTION, DATE_RESULT)}`,
  )}
  ${footnote(
    NOTE_WITHOUT_RESULT_ID,
    `<w:r><w:t xml:space="preserve">page </w:t></w:r>${complexField("PAGE", "")}` +
      `<w:r><w:t xml:space="preserve"> dated </w:t></w:r>${complexField(DATE_INSTRUCTION, "")}` +
      `<w:r><w:t xml:space="preserve"> end</w:t></w:r>`,
  )}
  ${footnote(NOTE_WITHOUT_FIELD_ID, `<w:r><w:t>${PLAIN_NOTE_TEXT}</w:t></w:r>`)}
  ${footnote(
    NOTE_WITH_CONTROL_ATOMS_ID,
    `<w:r><w:t>Left</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>Right</w:t></w:r>` +
      `<w:r><w:br/></w:r><w:r><w:t>Next</w:t></w:r>`,
  )}
  ${footnote(
    NOTE_WITH_STRUCTURED_FIELD_ID,
    `<w:r><w:t xml:space="preserve">clause </w:t></w:r>` +
      `<w:fldSimple w:instr=" PAGEREF _Ref1 \\h "><w:hyperlink w:anchor="_Ref1"><w:r><w:t>12</w:t></w:r></w:hyperlink></w:fldSimple>`,
  )}
</w:footnotes>`,
  );
  zip.file(
    "word/styles.xml",
    `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/**
 * Story text only comes from an editor document once the story has been loaded;
 * an untouched one is read straight off the parsed package. Load every header,
 * footer and note so the assertions cover the editor-document path.
 */
const openWithLoadedStories = async (buffer: ArrayBuffer): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  for (const { handle } of reviewer.listStories()) {
    if (handle.type !== "main") {
      reviewer.snapshotStory(handle);
    }
  }
  return reviewer;
};

/**
 * One surface reports every story, but each story kind reaches it through its
 * own reader, so an assertion names the lines of the kind it is about.
 */
const linesFor = (reviewer: FolioDocxReviewer, kind: string): string[] =>
  reviewer
    .getNotesAsText()
    .split("\n")
    .filter((line) => line.startsWith(`[${kind}`));

describe("field results in header, footer, and note text", () => {
  test("reads a note's fields as their stored results", async () => {
    const reviewer = await openWithLoadedStories(await createNoteFieldDocx());
    const lines = linesFor(reviewer, "footnote");

    expect(lines).toContain(`[footnote #${NOTE_WITH_RESULTS_ID}] page 7 dated 1 January 2020`);
    for (const line of lines) {
      expect(line).not.toContain("{page}");
      expect(line).not.toContain(new Date().toLocaleDateString());
    }
  });

  test("reads the same document the same way twice", async () => {
    const buffer = await createNoteFieldDocx();
    const reviewer = await openWithLoadedStories(buffer);
    const reopened = await openWithLoadedStories(buffer);

    expect(reviewer.getNotesAsText()).toBe(reviewer.getNotesAsText());
    expect(reopened.getNotesAsText()).toBe(reviewer.getNotesAsText());
    // Nothing that could have been synthesized from a clock: no d/m/y token in
    // any separator style, in any order.
    expect(reviewer.getNotesAsText()).not.toMatch(/\d{1,4}[/.-]\d{1,2}[/.-]\d{1,4}/u);
  });

  test("adds nothing for a note field with no stored result", async () => {
    const reviewer = await openWithLoadedStories(await createNoteFieldDocx());
    const lines = linesFor(reviewer, "footnote");

    expect(lines).toContain(`[footnote #${NOTE_WITHOUT_RESULT_ID}] page dated end`);
    for (const line of lines) {
      expect(line).not.toContain("{PAGE}");
      expect(line).not.toContain("{page}");
    }
  });

  test("leaves a note without a field byte-for-byte unchanged", async () => {
    const buffer = await createNoteFieldDocx();
    const loaded = await openWithLoadedStories(buffer);
    const unloaded = await FolioDocxReviewer.fromBuffer(buffer);
    const line = `[footnote #${NOTE_WITHOUT_FIELD_ID}] ${PLAIN_NOTE_TEXT}`;

    expect(loaded.getNotesAsText().split("\n")).toContain(line);
    expect(unloaded.getNotesAsText().split("\n")).toContain(line);
  });

  /**
   * Only the field branch is non-deterministic. A tab and a hard break carry no
   * text into a loaded story, and reading them as `\t` and `\n` would rewrite
   * every note holding one for anything that compares or hashes note text.
   */
  test("leaves a note's tabs and breaks contributing nothing", async () => {
    const reviewer = await openWithLoadedStories(await createNoteFieldDocx());

    expect(reviewer.getNotesAsText().split("\n")).toContain(
      `[footnote #${NOTE_WITH_CONTROL_ATOMS_ID}] LeftRightNext`,
    );
  });

  /**
   * Header and footer text shares this surface's output and this fix's helper,
   * so a header PAGE or DATE field is the same defect and needs its own
   * assertion: nothing else here would catch that one caller regressing.
   */
  test("reads a header's and a footer's fields as their stored results", async () => {
    const reviewer = await openWithLoadedStories(await createNoteFieldDocx());
    const lines = [...linesFor(reviewer, "header"), ...linesFor(reviewer, "footer")];

    expect(lines).toContain("[header default] header page 7 dated 1 January 2020 of end");
    expect(lines).toContain("[footer default] footer page 7 dated 1 January 2020 of end");
    for (const line of lines) {
      expect(line).not.toContain("{page}");
      expect(line).not.toContain(new Date().toLocaleDateString());
    }
  });

  test("reads a structured field from the runs that carry its result", async () => {
    const reviewer = await openWithLoadedStories(await createNoteFieldDocx());

    expect(reviewer.getNotesAsText().split("\n")).toContain(
      `[footnote #${NOTE_WITH_STRUCTURED_FIELD_ID}] clause 12`,
    );
  });
});
