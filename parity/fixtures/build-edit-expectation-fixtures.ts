#!/usr/bin/env bun
/** Build deterministic synthetic DOCX inputs for editing expectation work. */
import { mkdir } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

const FIXED_DATE = new Date("2026-01-01T00:00:00.000Z");
const OUTPUT_DIR = import.meta.dir;
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOCUMENT_REL = `${R_NS}/officeDocument`;
const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>`;
const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${DOCUMENT_REL}" Target="word/document.xml"/></Relationships>`;
const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${R_NS}/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="${R_NS}/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId4" Type="${R_NS}/endnotes" Target="endnotes.xml"/>
  <Relationship Id="rId5" Type="${R_NS}/comments" Target="comments.xml"/>
</Relationships>`;
const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="${W_NS}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/><w:sz w:val="30"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/></w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/></w:style>
  <w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
</w:styles>`;
const NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;
const FOOTNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="${W_NS}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Synthetic footnote text 42.</w:t></w:r></w:p></w:footnote></w:footnotes>`;
const ENDNOTES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:endnotes xmlns:w="${W_NS}"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="3"><w:p><w:pPr><w:pStyle w:val="EndnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> Synthetic endnote text 84.</w:t></w:r></w:p></w:endnote></w:endnotes>`;
const COMMENTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${W_NS}"><w:comment w:id="0" w:author="Synthetic Author" w:initials="SA" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Synthetic comment text 17.</w:t></w:r></w:p></w:comment></w:comments>`;

const paragraph = (text: string, properties = ""): string =>
  `<w:p>${properties}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const heading = (text: string): string =>
  paragraph(text, '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>');
const numbered = (text: string, level = 0): string =>
  paragraph(
    text,
    `<w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr></w:pPr>`,
  );
const cell = (text: string, merge = ""): string =>
  `<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/>${merge}</w:tcPr>${paragraph(text)}</w:tc>`;
const mergedTable = `<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid><w:tr>${cell("Alpha 31", '<w:vMerge w:val="restart"/>')}${cell("Beta 52")}</w:tr><w:tr>${cell("", "<w:vMerge/>")}${cell("Gamma 73")}</w:tr></w:tbl>`;
const sectionProperties =
  '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:num="2" w:space="720"/>';

const SEED_NAMES = [
  "edit-final-paragraph-seed.docx",
  "edit-paragraph-boundary-seed.docx",
  "edit-numbering-seed.docx",
  "edit-merged-table-seed.docx",
  "edit-comment-range-seed.docx",
  "edit-notes-fields-sections-seed.docx",
] as const;
const SEEDS = {
  "edit-final-paragraph-seed.docx": {
    body: `${heading("Final paragraph heading 11")}${paragraph("First paragraph Alpha 23.")}${paragraph("Target final paragraph Omega 47.")}`,
    references: "",
  },
  "edit-paragraph-boundary-seed.docx": {
    body: `${paragraph("Paragraph before boundary Cedar 19.")}${paragraph("Paragraph after boundary Maple 38.")}`,
    references: "",
  },
  "edit-numbering-seed.docx": {
    body: `${heading("Numbering seed 13")}${numbered("Primary numbered item 23.")}${numbered("Second primary numbered item 34.")}${numbered("Nested numbered item 37.", 1)}${numbered("Following primary item 41.")}`,
    references: "",
  },
  "edit-merged-table-seed.docx": {
    body: `${heading("Merged table seed 16")}${mergedTable}${paragraph("After table 79.")}`,
    references: "",
  },
  "edit-comment-range-seed.docx": {
    body: `${paragraph("Review span 18 begins here.", '<w:commentRangeStart w:id="0"/>').replace("</w:p>", '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>')}`,
    references: `<Relationship Id="rId5" Type="${R_NS}/comments" Target="comments.xml"/>`,
  },
  "edit-notes-fields-sections-seed.docx": {
    body: `${heading("Notes fields and sections seed 29")}<w:p><w:r><w:t xml:space="preserve">Footnote</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:t xml:space="preserve"> endnote</w:t></w:r><w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="3"/></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p><w:fldSimple w:instr=" MERGEFIELD CaseDate "><w:r><w:t>Cached field result 53.</w:t></w:r></w:fldSimple></w:p><w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/>${sectionProperties}</w:sectPr></w:pPr><w:r><w:t>Section boundary 59.</w:t></w:r></w:p>${heading("Second section 61.")}`,
    references: `<Relationship Id="rId3" Type="${R_NS}/footnotes" Target="footnotes.xml"/><Relationship Id="rId4" Type="${R_NS}/endnotes" Target="endnotes.xml"/>`,
  },
} as const;

export type EditExpectationSeedName = (typeof SEED_NAMES)[number];

const addPart = (zip: JSZip, name: string, contents: string): void => {
  zip.file(name, contents, { createFolders: false, date: FIXED_DATE });
};

const buildSeed = (name: EditExpectationSeedName): Promise<Uint8Array> => {
  const seed = SEEDS[name];
  const zip = new JSZip();
  addPart(zip, "[Content_Types].xml", CONTENT_TYPES);
  addPart(zip, "_rels/.rels", ROOT_RELS);
  addPart(
    zip,
    "word/_rels/document.xml.rels",
    DOCUMENT_RELS.replace(/\s*<Relationship Id="rId[345]"[^>]*\/>/g, "").replace(
      "</Relationships>",
      `${seed.references}</Relationships>`,
    ),
  );
  addPart(
    zip,
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${seed.body}<w:sectPr>${sectionProperties}</w:sectPr></w:body></w:document>`,
  );
  addPart(zip, "word/styles.xml", STYLES);
  addPart(zip, "word/numbering.xml", NUMBERING);
  if (name === "edit-comment-range-seed.docx") addPart(zip, "word/comments.xml", COMMENTS);
  if (name === "edit-notes-fields-sections-seed.docx") {
    addPart(zip, "word/footnotes.xml", FOOTNOTES);
    addPart(zip, "word/endnotes.xml", ENDNOTES);
  }
  return zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
};

export const buildEditExpectationSeeds = async (): Promise<
  ReadonlyMap<EditExpectationSeedName, Uint8Array>
> => {
  const result = new Map<EditExpectationSeedName, Uint8Array>();
  for (const name of SEED_NAMES) result.set(name, await buildSeed(name));
  return result;
};

const writeSeeds = async (): Promise<void> => {
  await mkdir(OUTPUT_DIR, { recursive: true });
  for (const [name, bytes] of await buildEditExpectationSeeds()) {
    await Bun.write(path.join(OUTPUT_DIR, name), bytes);
    console.log(`Wrote ${name} (${bytes.byteLength} bytes)`);
  }
};

const checkSeeds = async (): Promise<void> => {
  for (const [name, expected] of await buildEditExpectationSeeds()) {
    const actual = new Uint8Array(await Bun.file(path.join(OUTPUT_DIR, name)).arrayBuffer());
    if (!Bun.deepEquals(actual, expected))
      throw new TypeError(`${name} is stale; rebuild the synthetic edit seeds`);
  }
};

if (import.meta.main) await (process.argv.includes("--check") ? checkSeeds() : writeSeeds());
