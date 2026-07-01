/**
 * Round-trip tests for the numbering-definition write path.
 *
 * Parses a DOCX carrying a `word/numbering.xml` with two abstract numberings
 * (plus a `w:numPicBullet` and `w:nsid`/`w:tmpl` the model does NOT represent),
 * mutates one numbering definition in the model, saves through BOTH save paths,
 * re-parses, and asserts:
 * - the edited definition persists on save, and
 * - the UNEDITED definition plus the model-omitted parts (numPicBullet, the
 *   untouched abstractNum's nsid/tmpl) and every other package part stay
 *   byte-exact.
 *
 * Sibling of the footnote/endnote body write path (see noteSave.test.ts): a
 * save that changes only a body paragraph must leave numbering.xml verbatim.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import type { Document } from "../types/document";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx } from "./rezip";
import { attemptSelectiveSave } from "./selectiveSave";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

// Numbering ids used by the fixture. `EDITED_ABSTRACT_ID` is the definition the
// tests mutate; `UNTOUCHED_ABSTRACT_ID` must survive byte-exact.
const EDITED_ABSTRACT_ID = 0;
const UNTOUCHED_ABSTRACT_ID = 1;
const EDITED_NUM_ID = 1;

const ORIGINAL_LVL_TEXT = "%1.";
const EDITED_LVL_TEXT = "%1)";

// A picture-bullet definition and per-abstractNum nsid/tmpl — none of which the
// model represents. They must round-trip verbatim through both save paths.
const NUM_PIC_BULLET =
  '<w:numPicBullet w:numPicBulletId="0"><w:pict><v:shape id="_x0000_i1025" style="width:9pt;height:9pt"/></w:pict></w:numPicBullet>';
const UNTOUCHED_ABSTRACT =
  `<w:abstractNum w:abstractNumId="${UNTOUCHED_ABSTRACT_ID}" w15:restartNumberingAfterBreak="0">` +
  '<w:nsid w:val="1A2B3C4D"/><w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="0409000F"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>' +
  '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>';
// The edited abstractNum also carries nsid/tmpl the model omits; editing it
// re-emits it from the model (the documented safe-subset), so those are lost on
// the CHANGED definition only — the assertions below check the untouched one.
const EDITED_ABSTRACT =
  `<w:abstractNum w:abstractNumId="${EDITED_ABSTRACT_ID}" w15:restartNumberingAfterBreak="0">` +
  '<w:nsid w:val="0F1E2D3C"/><w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="04090001"/>' +
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="${ORIGINAL_LVL_TEXT}"/>` +
  '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>';

const numberingXml =
  `${XML_DECLARATION}\n` +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml">' +
  NUM_PIC_BULLET +
  EDITED_ABSTRACT +
  UNTOUCHED_ABSTRACT +
  `<w:num w:numId="${EDITED_NUM_ID}"><w:abstractNumId w:val="${EDITED_ABSTRACT_ID}"/></w:num>` +
  `<w:num w:numId="2"><w:abstractNumId w:val="${UNTOUCHED_ABSTRACT_ID}"/></w:num>` +
  "</w:numbering>";

const BODY_PARA_ID = "B0000001";
const documentXml = `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="${BODY_PARA_ID}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${EDITED_NUM_ID}"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">First item</w:t></w:r></w:p>
    <w:p w14:paraId="B0000002"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">Bullet item</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;

const stylesXml = `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
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
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`;

const packageRelsXml = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>`;

const documentRelsXml = `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.numbering}" Target="numbering.xml"/>
</Relationships>`;

async function buildDocx(parts: {
  numberingXml: string;
  documentXml: string;
}): Promise<ArrayBuffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", contentTypesXml);
  zip.file("_rels/.rels", packageRelsXml);
  zip.file("word/_rels/document.xml.rels", documentRelsXml);
  zip.file("word/document.xml", parts.documentXml);
  zip.file("word/styles.xml", stylesXml);
  zip.file("word/numbering.xml", parts.numberingXml);
  zip.file("docProps/core.xml", corePropertiesXml);
  return zip.generateAsync({ type: "arraybuffer" });
}

async function createNumberingFixture(): Promise<ArrayBuffer> {
  return buildDocx({ numberingXml, documentXml });
}

async function readPart(buffer: ArrayBuffer, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(path);
  if (!file) {
    throw new Error(`No ${path} in package`);
  }
  return file.async("text");
}

function editFirstLevelText(doc: Document, abstractNumId: number, newText: string): void {
  const abstractNum = doc.package.numbering?.abstractNums.find(
    (a) => a.abstractNumId === abstractNumId,
  );
  const level = abstractNum?.levels.find((l) => l.ilvl === 0);
  if (!level) {
    throw new Error(`expected level 0 on abstractNum ${abstractNumId}`);
  }
  level.lvlText = newText;
}

function levelText(doc: Document, abstractNumId: number): string | undefined {
  return doc.package.numbering?.abstractNums
    .find((a) => a.abstractNumId === abstractNumId)
    ?.levels.find((l) => l.ilvl === 0)?.lvlText;
}

describe("numbering-definition write path (selective save)", () => {
  test("edited numbering definition persists; untouched definition + omitted parts stay byte-exact", async () => {
    const buffer = await createNumberingFixture();
    const originalDocumentXml = await readPart(buffer, "word/document.xml");
    const originalStylesXml = await readPart(buffer, "word/styles.xml");

    const doc = await parseDocx(buffer, { preloadFonts: false });
    editFirstLevelText(doc, EDITED_ABSTRACT_ID, EDITED_LVL_TEXT);

    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });

    expect(result).not.toBeNull();
    if (!result) {
      throw new Error("selective save returned null");
    }

    // The edit round-trips through re-parse.
    const reparsed = await parseDocx(result, { preloadFonts: false });
    expect(levelText(reparsed, EDITED_ABSTRACT_ID)).toBe(EDITED_LVL_TEXT);

    // Other parts are untouched.
    expect(await readPart(result, "word/document.xml")).toBe(originalDocumentXml);
    expect(await readPart(result, "word/styles.xml")).toBe(originalStylesXml);

    // Within numbering.xml: the edit landed, the old marker is gone, and the
    // untouched definition + the model-omitted picture bullet stay byte-exact.
    const savedNumberingXml = await readPart(result, "word/numbering.xml");
    expect(savedNumberingXml).toContain(`<w:lvlText w:val="${EDITED_LVL_TEXT}"/>`);
    expect(savedNumberingXml).not.toContain(`<w:lvlText w:val="${ORIGINAL_LVL_TEXT}"/>`);
    expect(savedNumberingXml).toContain(UNTOUCHED_ABSTRACT);
    expect(savedNumberingXml).toContain(NUM_PIC_BULLET);
  });

  test("a body-only edit leaves numbering.xml byte-exact", async () => {
    const buffer = await createNumberingFixture();
    const originalNumberingXml = await readPart(buffer, "word/numbering.xml");

    const doc = await parseDocx(buffer, { preloadFonts: false });
    const bodyPara = doc.package.document.content[0];
    if (!bodyPara || bodyPara.type !== "paragraph") {
      throw new Error("expected a body paragraph");
    }
    for (const item of bodyPara.content) {
      if (item.type === "run") {
        for (const runContent of item.content) {
          if (runContent.type === "text") {
            runContent.text = "First item edited";
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

    // numbering.xml is never touched by a body-only edit.
    expect(await readPart(result, "word/numbering.xml")).toBe(originalNumberingXml);
    expect(await readPart(result, "word/document.xml")).toContain("First item edited");
  });
});

describe("numbering-definition write path (full repack)", () => {
  test("edited numbering definition persists through repackDocx; untouched parts intact", async () => {
    const buffer = await createNumberingFixture();

    const doc = await parseDocx(buffer, { preloadFonts: false });
    editFirstLevelText(doc, EDITED_ABSTRACT_ID, EDITED_LVL_TEXT);

    const result = await repackDocx(doc);

    const reparsed = await parseDocx(result, { preloadFonts: false });
    expect(levelText(reparsed, EDITED_ABSTRACT_ID)).toBe(EDITED_LVL_TEXT);

    const savedNumberingXml = await readPart(result, "word/numbering.xml");
    expect(savedNumberingXml).toContain(`<w:lvlText w:val="${EDITED_LVL_TEXT}"/>`);
    expect(savedNumberingXml).not.toContain(`<w:lvlText w:val="${ORIGINAL_LVL_TEXT}"/>`);
    expect(savedNumberingXml).toContain(UNTOUCHED_ABSTRACT);
    expect(savedNumberingXml).toContain(NUM_PIC_BULLET);
  });

  test("a repack with no numbering edit leaves numbering.xml byte-exact", async () => {
    const buffer = await createNumberingFixture();
    const originalNumberingXml = await readPart(buffer, "word/numbering.xml");

    const doc = await parseDocx(buffer, { preloadFonts: false });
    const result = await repackDocx(doc);

    // Byte-exact preservation includes the parts the model does not represent:
    // an unedited numbering is never re-serialized, so numPicBullet, nsid, and
    // tmpl all survive even though the model drops them.
    const savedNumberingXml = await readPart(result, "word/numbering.xml");
    expect(savedNumberingXml).toBe(originalNumberingXml);
    expect(savedNumberingXml).toContain(NUM_PIC_BULLET);
    expect(savedNumberingXml).toContain('<w:tmpl w:val="04090001"/>');

    // The definitions still round-trip unchanged.
    const reparsed = await parseDocx(result, { preloadFonts: false });
    expect(levelText(reparsed, EDITED_ABSTRACT_ID)).toBe(ORIGINAL_LVL_TEXT);
  });

  test("selective and full repack produce the same numbering.xml after an edit", async () => {
    const buffer = await createNumberingFixture();

    const docForSelective = await parseDocx(buffer, { preloadFonts: false });
    editFirstLevelText(docForSelective, EDITED_ABSTRACT_ID, EDITED_LVL_TEXT);
    const selective = await attemptSelectiveSave(docForSelective, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(selective).not.toBeNull();
    if (!selective) {
      throw new Error("selective save returned null");
    }

    const docForFull = await parseDocx(buffer, { preloadFonts: false });
    editFirstLevelText(docForFull, EDITED_ABSTRACT_ID, EDITED_LVL_TEXT);
    const full = await repackDocx({ ...docForFull, originalBuffer: buffer });

    expect(await readPart(selective, "word/numbering.xml")).toBe(
      await readPart(full, "word/numbering.xml"),
    );
  });
});

// A document referencing a single numbering instance, for the focused fixtures
// below.
const singleListDocumentXml = (numId: number) => `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p w14:paraId="C0000001"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">Item</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
  </w:body>
</w:document>`;

describe("numbering-definition write path (newline-formatted opening tags)", () => {
  // The abstractNum / num opening tags carry a newline before their id
  // attribute, which is valid XML. The element scanner must treat it as a name
  // boundary (not a longer-named sibling) or the patch silently no-ops.
  const newlineNumberingXml =
    `${XML_DECLARATION}\n` +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum\n    w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>' +
    `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="${ORIGINAL_LVL_TEXT}"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>` +
    '<w:num\n    w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    "</w:numbering>";

  test("edit persists through both save paths despite newline-formatted tags", async () => {
    const buffer = await buildDocx({
      numberingXml: newlineNumberingXml,
      documentXml: singleListDocumentXml(1),
    });

    const selectiveDoc = await parseDocx(buffer, { preloadFonts: false });
    editFirstLevelText(selectiveDoc, 0, EDITED_LVL_TEXT);
    const selective = await attemptSelectiveSave(selectiveDoc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(selective).not.toBeNull();
    if (!selective) {
      throw new Error("selective save returned null");
    }
    expect(levelText(await parseDocx(selective, { preloadFonts: false }), 0)).toBe(EDITED_LVL_TEXT);
    expect(await readPart(selective, "word/numbering.xml")).toContain(
      `<w:lvlText w:val="${EDITED_LVL_TEXT}"/>`,
    );

    const fullDoc = await parseDocx(buffer, { preloadFonts: false });
    editFirstLevelText(fullDoc, 0, EDITED_LVL_TEXT);
    const full = await repackDocx({ ...fullDoc, originalBuffer: buffer });
    expect(levelText(await parseDocx(full, { preloadFonts: false }), 0)).toBe(EDITED_LVL_TEXT);
  });
});

describe("numbering-definition write path (custom / AlternateContent numFmt)", () => {
  // abstractNum 0 level 0 carries a Word custom number format wrapped in
  // mc:AlternateContent. folio's model flattens it to a synthetic decimalZero4
  // it cannot re-emit as valid OOXML, so editing an UNRELATED field (lvlText)
  // must preserve the original numFmt element rather than write the synthetic.
  const CUSTOM_ALT_CONTENT =
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
    '<mc:Choice Requires="w14"><w:numFmt w:val="custom" w:format="0001, 0002, 0003"/></mc:Choice>' +
    '<mc:Fallback><w:numFmt w:val="decimal"/></mc:Fallback></mc:AlternateContent>';
  const customNumberingXml =
    `${XML_DECLARATION}\n` +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
    '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>' +
    `<w:lvl w:ilvl="0"><w:start w:val="1"/>${CUSTOM_ALT_CONTENT}<w:lvlText w:val="${ORIGINAL_LVL_TEXT}"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>` +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    "</w:numbering>";

  const expectCustomFormatSurvives = async (saved: ArrayBuffer): Promise<void> => {
    const savedNumberingXml = await readPart(saved, "word/numbering.xml");
    // The original custom numFmt element is preserved verbatim...
    expect(savedNumberingXml).toContain(CUSTOM_ALT_CONTENT);
    // ...and no synthetic (non-OOXML) value leaked into the saved file.
    expect(savedNumberingXml).not.toContain("decimalZero4");
    // The edit landed and the format still round-trips to the synthetic model value.
    expect(savedNumberingXml).toContain(`<w:lvlText w:val="${EDITED_LVL_TEXT}"/>`);
    const reparsed = await parseDocx(saved, { preloadFonts: false });
    expect(levelText(reparsed, 0)).toBe(EDITED_LVL_TEXT);
    const level = reparsed.package.numbering?.abstractNums
      .find((a) => a.abstractNumId === 0)
      ?.levels.find((l) => l.ilvl === 0);
    expect(level?.numFmt).toBe("decimalZero4");
  };

  test("selective save preserves the custom format when a different field is edited", async () => {
    const buffer = await buildDocx({
      numberingXml: customNumberingXml,
      documentXml: singleListDocumentXml(1),
    });
    // Sanity: the model flattened the custom format to the synthetic value.
    const doc = await parseDocx(buffer, { preloadFonts: false });
    const level = doc.package.numbering?.abstractNums
      .find((a) => a.abstractNumId === 0)
      ?.levels.find((l) => l.ilvl === 0);
    expect(level?.numFmt).toBe("decimalZero4");

    editFirstLevelText(doc, 0, EDITED_LVL_TEXT);
    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(result).not.toBeNull();
    if (!result) {
      throw new Error("selective save returned null");
    }
    await expectCustomFormatSurvives(result);
  });

  test("full repack preserves the custom format when a different field is edited", async () => {
    const buffer = await buildDocx({
      numberingXml: customNumberingXml,
      documentXml: singleListDocumentXml(1),
    });
    const doc = await parseDocx(buffer, { preloadFonts: false });
    editFirstLevelText(doc, 0, EDITED_LVL_TEXT);
    const result = await repackDocx({ ...doc, originalBuffer: buffer });
    await expectCustomFormatSurvives(result);
  });

  // A genuine format edit (decimalZero4 -> decimalZero5) must NOT be reverted to
  // the original element, and must never emit the non-OOXML synthetic literal.
  const editLevelNumFmt = (doc: Document, numFmt: "decimalZero5"): void => {
    const level = doc.package.numbering?.abstractNums
      .find((a) => a.abstractNumId === 0)
      ?.levels.find((l) => l.ilvl === 0);
    if (!level) {
      throw new Error("expected level 0 on abstractNum 0");
    }
    level.numFmt = numFmt;
  };

  const expectWidthChangedToFive = async (saved: ArrayBuffer): Promise<void> => {
    const savedNumberingXml = await readPart(saved, "word/numbering.xml");
    // The NEW width (5) is emitted as a valid OOXML custom format...
    expect(savedNumberingXml).toContain('<w:numFmt w:val="custom" w:format="00001"/>');
    // ...not the synthetic literal, and not the original width-4 format.
    expect(savedNumberingXml).not.toContain("decimalZero5");
    expect(savedNumberingXml).not.toContain("0001, 0002, 0003");
    // It round-trips to the edited synthetic value.
    const reparsed = await parseDocx(saved, { preloadFonts: false });
    const level = reparsed.package.numbering?.abstractNums
      .find((a) => a.abstractNumId === 0)
      ?.levels.find((l) => l.ilvl === 0);
    expect(level?.numFmt).toBe("decimalZero5");
  };

  test("selective save honors an intentional synthetic width change (4 -> 5)", async () => {
    const buffer = await buildDocx({
      numberingXml: customNumberingXml,
      documentXml: singleListDocumentXml(1),
    });
    const doc = await parseDocx(buffer, { preloadFonts: false });
    editLevelNumFmt(doc, "decimalZero5");
    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(result).not.toBeNull();
    if (!result) {
      throw new Error("selective save returned null");
    }
    await expectWidthChangedToFive(result);
  });

  test("full repack honors an intentional synthetic width change (4 -> 5)", async () => {
    const buffer = await buildDocx({
      numberingXml: customNumberingXml,
      documentXml: singleListDocumentXml(1),
    });
    const doc = await parseDocx(buffer, { preloadFonts: false });
    editLevelNumFmt(doc, "decimalZero5");
    const result = await repackDocx({ ...doc, originalBuffer: buffer });
    await expectWidthChangedToFive(result);
  });
});

describe("numbering-definition write path (ancestor-scoped xmlns on restored numFmt)", () => {
  // The custom format's `mc:` prefix (and the `w14` its Choice requires) are
  // declared on the w:abstractNum ancestor, NOT the numbering root. Re-emitting
  // the definition drops those ancestor declarations, so the restored numFmt
  // element must carry them or the saved part has an undefined prefix.
  const MC_URI = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  const W14_URI = "http://schemas.microsoft.com/office/word/2010/wordml";
  const SCOPED_ALT_CONTENT =
    "<mc:AlternateContent>" +
    '<mc:Choice Requires="w14"><w:numFmt w:val="custom" w:format="0001, 0002, 0003"/></mc:Choice>' +
    '<mc:Fallback><w:numFmt w:val="decimal"/></mc:Fallback></mc:AlternateContent>';
  const scopedNumberingXml =
    `${XML_DECLARATION}\n` +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:abstractNum w:abstractNumId="0" xmlns:mc="${MC_URI}" xmlns:w14="${W14_URI}">` +
    '<w:multiLevelType w:val="hybridMultilevel"/>' +
    `<w:lvl w:ilvl="0"><w:start w:val="1"/>${SCOPED_ALT_CONTENT}<w:lvlText w:val="${ORIGINAL_LVL_TEXT}"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>` +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
    "</w:numbering>";

  const expectSelfContainedFormat = async (saved: ArrayBuffer): Promise<void> => {
    const savedNumberingXml = await readPart(saved, "word/numbering.xml");
    // The restored numFmt element now declares the prefixes its ancestor used,
    // so no dangling mc:/w14 prefix survives in the saved part.
    expect(savedNumberingXml).toContain(`<mc:AlternateContent xmlns:mc="${MC_URI}"`);
    expect(savedNumberingXml).toContain(`xmlns:w14="${W14_URI}"`);
    expect(savedNumberingXml).not.toContain("decimalZero4");
    expect(savedNumberingXml).toContain(`<w:lvlText w:val="${EDITED_LVL_TEXT}"/>`);
    const reparsed = await parseDocx(saved, { preloadFonts: false });
    expect(levelText(reparsed, 0)).toBe(EDITED_LVL_TEXT);
    const level = reparsed.package.numbering?.abstractNums
      .find((a) => a.abstractNumId === 0)
      ?.levels.find((l) => l.ilvl === 0);
    expect(level?.numFmt).toBe("decimalZero4");
  };

  test("selective save makes the restored custom format self-contained", async () => {
    const buffer = await buildDocx({
      numberingXml: scopedNumberingXml,
      documentXml: singleListDocumentXml(1),
    });
    const doc = await parseDocx(buffer, { preloadFonts: false });
    editFirstLevelText(doc, 0, EDITED_LVL_TEXT);
    const result = await attemptSelectiveSave(doc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(result).not.toBeNull();
    if (!result) {
      throw new Error("selective save returned null");
    }
    await expectSelfContainedFormat(result);
  });

  test("full repack makes the restored custom format self-contained", async () => {
    const buffer = await buildDocx({
      numberingXml: scopedNumberingXml,
      documentXml: singleListDocumentXml(1),
    });
    const doc = await parseDocx(buffer, { preloadFonts: false });
    editFirstLevelText(doc, 0, EDITED_LVL_TEXT);
    const result = await repackDocx({ ...doc, originalBuffer: buffer });
    await expectSelfContainedFormat(result);
  });
});

describe("numbering-definition write path (non-w: prefix graceful degradation)", () => {
  // numbering.xml binds the WordprocessingML namespace to a non-conventional
  // prefix (`wp:`). The splice matches definitions by the literal `w:` prefix,
  // so it finds nothing to splice: the edit is not applied, but the part is left
  // byte-exact verbatim (no corruption / malformed output on either save path).
  // Real Word always uses `w:`; this only guards the pathological input.
  const WML_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const altPrefixNumberingXml =
    `${XML_DECLARATION}\n` +
    `<wp:numbering xmlns:wp="${WML_URI}">` +
    '<wp:abstractNum wp:abstractNumId="0"><wp:multiLevelType wp:val="hybridMultilevel"/>' +
    `<wp:lvl wp:ilvl="0"><wp:start wp:val="1"/><wp:numFmt wp:val="decimal"/><wp:lvlText wp:val="${ORIGINAL_LVL_TEXT}"/><wp:lvlJc wp:val="left"/></wp:lvl></wp:abstractNum>` +
    '<wp:num wp:numId="1"><wp:abstractNumId wp:val="0"/></wp:num>' +
    "</wp:numbering>";

  const expectNumberingByteExact = async (saved: ArrayBuffer, original: string): Promise<void> => {
    // The alt-prefix part is left verbatim (edit not applied, nothing corrupted).
    expect(await readPart(saved, "word/numbering.xml")).toBe(original);
    // And the saved document still parses.
    await expect(parseDocx(saved, { preloadFonts: false })).resolves.toBeDefined();
  };

  // The parser resolves element/attribute NAMES by local name, so the abstract
  // numbering still populates from the `wp:` part; change a modeled field so the
  // write path attempts a splice, then finds no `<w:abstractNum` to match.
  const forceModelChange = (doc: Document): void => {
    const abstractNum = doc.package.numbering?.abstractNums[0];
    if (!abstractNum) {
      throw new Error("expected the parser to populate at least one abstractNum");
    }
    abstractNum.multiLevelType = "multilevel";
  };

  test("edit degrades to a byte-exact no-op on both save paths", async () => {
    const buffer = await buildDocx({
      numberingXml: altPrefixNumberingXml,
      documentXml: singleListDocumentXml(1),
    });
    const originalNumberingXml = await readPart(buffer, "word/numbering.xml");

    const selectiveDoc = await parseDocx(buffer, { preloadFonts: false });
    forceModelChange(selectiveDoc);
    const selective = await attemptSelectiveSave(selectiveDoc, buffer, {
      changedParaIds: new Set(),
      structuralChange: false,
      hasUntrackedChanges: false,
    });
    expect(selective).not.toBeNull();
    if (!selective) {
      throw new Error("selective save returned null");
    }
    await expectNumberingByteExact(selective, originalNumberingXml);

    const fullDoc = await parseDocx(buffer, { preloadFonts: false });
    forceModelChange(fullDoc);
    const full = await repackDocx({ ...fullDoc, originalBuffer: buffer });
    await expectNumberingByteExact(full, originalNumberingXml);
  });
});
