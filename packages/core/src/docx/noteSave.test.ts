/**
 * Round-trip tests for the footnote/endnote body write path.
 *
 * Parses a DOCX carrying a footnote and an endnote (each with the separator
 * notes Word requires), mutates a note body in the model, runs the selective
 * save, re-parses, and asserts:
 * - the edited note text persists on save, and
 * - the UNEDITED note part plus every other part (document.xml, styles.xml,
 *   the separator notes) stay byte-exact.
 *
 * Regression guard: a save that changes only a body paragraph must leave both
 * note parts verbatim (the write path only fires when a note body changed).
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import type { Endnote, Footnote } from "../types/document";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx } from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

// A footnotes part with the two separator notes Word always emits (ids -1 / 0)
// followed by one normal note whose paragraph carries a paraId so the selective
// patch can target it. `ORIGINAL_FOOTNOTE_TEXT` is the editable body text.
const ORIGINAL_FOOTNOTE_TEXT = "Original footnote body";
const FOOTNOTE_PARA_ID = "F1000001";
const FOOTNOTE_SEPARATORS =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';

// The normal note carries the in-note `w:footnoteRef` auto-number mark that
// real Word documents emit and that the model does NOT represent, so the
// byte-exact assertions below confirm an unedited note keeps that mark.
const footnotesXml = `${XML_DECLARATION}
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${FOOTNOTE_SEPARATORS}<w:footnote w:id="1"><w:p w14:paraId="${FOOTNOTE_PARA_ID}"><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve">${ORIGINAL_FOOTNOTE_TEXT}</w:t></w:r></w:p></w:footnote></w:footnotes>`;

const ORIGINAL_ENDNOTE_TEXT = "Original endnote body";
const ENDNOTE_PARA_ID = "E1000001";
const ENDNOTE_SEPARATORS =
  '<w:endnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:endnote>' +
  '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>';

const endnotesXml = `${XML_DECLARATION}
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">${ENDNOTE_SEPARATORS}<w:endnote w:id="1"><w:p w14:paraId="${ENDNOTE_PARA_ID}"><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve">${ORIGINAL_ENDNOTE_TEXT}</w:t></w:r></w:p></w:endnote></w:endnotes>`;

const BODY_PARA_ID = "B0000001";
const documentXml = `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="${BODY_PARA_ID}"><w:r><w:t xml:space="preserve">Body text</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;

const stylesXml = `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/></w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/></w:style>
</w:styles>`;

const corePropertiesXml = `${XML_DECLARATION}
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00.000Z</dcterms:modified>
</cp:coreProperties>`;

const contentTypesXml = `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const packageRelsXml = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const documentRelsXml = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.footnotes}" Target="footnotes.xml"/>
  <Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.endnotes}" Target="endnotes.xml"/>
</Relationships>`;

async function createNotesFixture(): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", packageRelsXml);
  zip.file("word/_rels/document.xml.rels", documentRelsXml);
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/footnotes.xml", footnotesXml);
  zip.file("word/endnotes.xml", endnotesXml);
  zip.file("docProps/core.xml", corePropertiesXml);
  return zip.generateAsync({ type: "arraybuffer" });
}

async function readPart(buffer: ArrayBuffer, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(path);
  if (!file) {
    throw new Error(`No ${path} in package`);
  }
  return file.async("text");
}

function setFirstNoteText(note: Footnote | Endnote, newText: string): void {
  const block = note.content[0];
  if (!block || block.type !== "paragraph") {
    throw new Error("expected a paragraph as the note's first block");
  }
  for (const item of block.content) {
    if (item.type !== "run") {
      continue;
    }
    for (const runContent of item.content) {
      if (runContent.type === "text") {
        runContent.text = newText;
        return;
      }
    }
  }
  throw new Error("expected a text run in the note paragraph");
}

function noteBodyText(note: Footnote | Endnote | undefined): string {
  return (note?.content ?? [])
    .flatMap((block) => (block.type === "paragraph" ? block.content : []))
    .flatMap((item) => (item.type === "run" ? item.content : []))
    .map((rc) => (rc.type === "text" ? rc.text : ""))
    .join("");
}

describe("footnote / endnote body write path (selective save)", () => {
  test("edited footnote body persists; endnote and other parts stay byte-exact", async () => {
    const buffer = await createNotesFixture();
    const originalEndnotesXml = await readPart(buffer, "word/endnotes.xml");
    const originalDocumentXml = await readPart(buffer, "word/document.xml");
    const originalStylesXml = await readPart(buffer, "word/styles.xml");

    const doc = await parseDocx(buffer, { preloadFonts: false });
    expect(doc.package.footnotes?.length).toBe(1);

    const footnote = doc.package.footnotes?.[0];
    if (!footnote) {
      throw new Error("expected a parsed footnote");
    }
    const editedText = "Edited footnote body";
    setFirstNoteText(footnote, editedText);

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set([FOOTNOTE_PARA_ID]),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).not.toBeNull();
    if (!result) {
      throw new Error("selective save returned null");
    }

    // The edited footnote text round-trips through re-parse.
    const reparsed = await parseDocx(result, { preloadFonts: false });
    expect(noteBodyText(reparsed.package.footnotes?.[0])).toBe(editedText);

    // The unedited endnote part and other parts are untouched.
    expect(await readPart(result, "word/endnotes.xml")).toBe(originalEndnotesXml);
    expect(await readPart(result, "word/document.xml")).toBe(originalDocumentXml);
    expect(await readPart(result, "word/styles.xml")).toBe(originalStylesXml);

    // Within footnotes.xml, the required separator notes stay byte-exact and
    // the old body text is gone, replaced by the edit.
    const savedFootnotesXml = await readPart(result, "word/footnotes.xml");
    expect(savedFootnotesXml).toContain(FOOTNOTE_SEPARATORS);
    expect(savedFootnotesXml).toContain(editedText);
    expect(savedFootnotesXml).not.toContain(ORIGINAL_FOOTNOTE_TEXT);
  });

  test("edited endnote body persists; footnote and other parts stay byte-exact", async () => {
    const buffer = await createNotesFixture();
    const originalFootnotesXml = await readPart(buffer, "word/footnotes.xml");

    const doc = await parseDocx(buffer, { preloadFonts: false });
    const endnote = doc.package.endnotes?.[0];
    if (!endnote) {
      throw new Error("expected a parsed endnote");
    }
    const editedText = "Edited endnote body";
    setFirstNoteText(endnote, editedText);

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set([ENDNOTE_PARA_ID]),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).not.toBeNull();
    if (!result) {
      throw new Error("selective save returned null");
    }

    const reparsed = await parseDocx(result, { preloadFonts: false });
    expect(noteBodyText(reparsed.package.endnotes?.[0])).toBe(editedText);

    // The footnote part is untouched when only an endnote changed.
    expect(await readPart(result, "word/footnotes.xml")).toBe(originalFootnotesXml);

    const savedEndnotesXml = await readPart(result, "word/endnotes.xml");
    expect(savedEndnotesXml).toContain(ENDNOTE_SEPARATORS);
    expect(savedEndnotesXml).toContain(editedText);
    expect(savedEndnotesXml).not.toContain(ORIGINAL_ENDNOTE_TEXT);
  });

  test("a body-only edit leaves both note parts byte-exact", async () => {
    const buffer = await createNotesFixture();
    const originalFootnotesXml = await readPart(buffer, "word/footnotes.xml");
    const originalEndnotesXml = await readPart(buffer, "word/endnotes.xml");

    const doc = await parseDocx(buffer, { preloadFonts: false });
    const bodyPara = doc.package.document.content[0];
    if (!bodyPara || bodyPara.type !== "paragraph") {
      throw new Error("expected a body paragraph");
    }
    for (const item of bodyPara.content) {
      if (item.type === "run") {
        for (const runContent of item.content) {
          if (runContent.type === "text") {
            runContent.text = "Body text edited";
          }
        }
      }
    }

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set([BODY_PARA_ID]),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).not.toBeNull();
    if (!result) {
      throw new Error("selective save returned null");
    }

    // Note parts are never touched by a body-only edit.
    expect(await readPart(result, "word/footnotes.xml")).toBe(originalFootnotesXml);
    expect(await readPart(result, "word/endnotes.xml")).toBe(originalEndnotesXml);

    // The body edit landed in document.xml.
    expect(await readPart(result, "word/document.xml")).toContain("Body text edited");
  });
});

describe("footnote / endnote body write path (full repack)", () => {
  test("edited footnote body persists through repackDocx; endnote + separators intact", async () => {
    const buffer = await createNotesFixture();
    const originalEndnotesXml = await readPart(buffer, "word/endnotes.xml");

    const doc = await parseDocx(buffer, { preloadFonts: false });
    const footnote = doc.package.footnotes?.[0];
    if (!footnote) {
      throw new Error("expected a parsed footnote");
    }
    const editedText = "Edited footnote body via full repack";
    setFirstNoteText(footnote, editedText);

    const result = await repackDocx(doc);

    const reparsed = await parseDocx(result, { preloadFonts: false });
    expect(noteBodyText(reparsed.package.footnotes?.[0])).toBe(editedText);

    // Full repack rebuilds document.xml from the model, but the unedited endnote
    // part is left byte-exact (no endnote paragraph changed).
    expect(await readPart(result, "word/endnotes.xml")).toBe(originalEndnotesXml);

    // Separator notes stay byte-exact inside footnotes.xml; the old body text is
    // gone, replaced by the edit.
    const savedFootnotesXml = await readPart(result, "word/footnotes.xml");
    expect(savedFootnotesXml).toContain(FOOTNOTE_SEPARATORS);
    expect(savedFootnotesXml).toContain(editedText);
    expect(savedFootnotesXml).not.toContain(ORIGINAL_FOOTNOTE_TEXT);
  });

  test("edited endnote body persists through repackDocx; footnote + separators intact", async () => {
    const buffer = await createNotesFixture();
    const originalFootnotesXml = await readPart(buffer, "word/footnotes.xml");

    const doc = await parseDocx(buffer, { preloadFonts: false });
    const endnote = doc.package.endnotes?.[0];
    if (!endnote) {
      throw new Error("expected a parsed endnote");
    }
    const editedText = "Edited endnote body via full repack";
    setFirstNoteText(endnote, editedText);

    const result = await repackDocx(doc);

    const reparsed = await parseDocx(result, { preloadFonts: false });
    expect(noteBodyText(reparsed.package.endnotes?.[0])).toBe(editedText);

    // The unedited footnote part is left byte-exact.
    expect(await readPart(result, "word/footnotes.xml")).toBe(originalFootnotesXml);

    const savedEndnotesXml = await readPart(result, "word/endnotes.xml");
    expect(savedEndnotesXml).toContain(ENDNOTE_SEPARATORS);
    expect(savedEndnotesXml).toContain(editedText);
    expect(savedEndnotesXml).not.toContain(ORIGINAL_ENDNOTE_TEXT);
  });

  test("a repack with no note edit leaves both note parts byte-exact", async () => {
    const buffer = await createNotesFixture();
    const originalFootnotesXml = await readPart(buffer, "word/footnotes.xml");
    const originalEndnotesXml = await readPart(buffer, "word/endnotes.xml");

    const doc = await parseDocx(buffer, { preloadFonts: false });
    const result = await repackDocx(doc);

    // Byte-exact preservation includes the in-note auto-number marks the model
    // does not represent: an unedited note is never re-serialized, so nothing is
    // lost even though `w:footnoteRef` / `w:endnoteRef` are absent from the model.
    const savedFootnotesXml = await readPart(result, "word/footnotes.xml");
    const savedEndnotesXml = await readPart(result, "word/endnotes.xml");
    expect(savedFootnotesXml).toBe(originalFootnotesXml);
    expect(savedEndnotesXml).toBe(originalEndnotesXml);
    expect(savedFootnotesXml).toContain("<w:footnoteRef/>");
    expect(savedEndnotesXml).toContain("<w:endnoteRef/>");

    // The note bodies still round-trip unchanged.
    const reparsed = await parseDocx(result, { preloadFonts: false });
    expect(noteBodyText(reparsed.package.footnotes?.[0])).toBe(ORIGINAL_FOOTNOTE_TEXT);
    expect(noteBodyText(reparsed.package.endnotes?.[0])).toBe(ORIGINAL_ENDNOTE_TEXT);
  });
});
