#!/usr/bin/env bun
/**
 * Build the synthetic DOCX that exercises cursive joining across run
 * boundaries.
 *
 * Generated from hand-written OOXML so the runs stay reviewable: what matters
 * is exactly where each word is split and which property differs across the
 * split, and that is invisible in a binary fixture.
 *
 * Three paragraphs, each a different answer to "should this boundary need
 * repair?":
 *
 *  1. bold on the second half of a word  -> different face, joining severed
 *  2. colour on the second half          -> same face, browser joins already
 *  3. Latin with the same splits         -> no cursive letters, never touched
 *
 * Run: bun run build:fixture:cursive-face-change
 */
import { mkdir } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

const OUTPUT_DIR = import.meta.dir;
const OUTPUT_PATH = path.join(OUTPUT_DIR, "cursive-face-change.docx");
/** Fixed so the fixture is byte-reproducible across rebuilds. */
const FIXED_DATE = new Date("2026-01-01T00:00:00.000Z");

// Arabic, authored as escapes rather than pasted glyphs: a pasted glyph once
// silently corrupted a character class in this repo. Each word below is fully
// connected, so every internal split severs a join.
const MEEM = "م";
const KAF = "ك";
const TEH = "ت";
const BEH = "ب";
const ALEF = "ا";
const LAM = "ل";
const HAH = "ح";
const MEEM2 = "م";
const QAF = "ق";
const YEH = "ي";
const NOON = "ن";
const SEEN = "س";
const DAL = "د";
const WAW = "و";
const REH = "ر";
const TEH_MARBUTA = "ة";

/** "مكتب" (office) — split after 2 letters lands mid-word. */
const OFFICE_HEAD = MEEM + KAF;
const OFFICE_TAIL = TEH + BEH;
/** "المحكمة" (the court) — split inside the connected middle. */
const COURT_HEAD = ALEF + LAM + MEEM2 + HAH;
const COURT_TAIL = KAF + MEEM2 + TEH_MARBUTA;
/** "الدستورية" (constitutional) — a longer connected span to split. */
const CONSTITUTIONAL_HEAD = ALEF + LAM + DAL + SEEN;
const CONSTITUTIONAL_TAIL = TEH + WAW + REH + YEH + TEH_MARBUTA;
/** "قانون" (law) — split after qaf. */
const LAW_HEAD = QAF + ALEF;
const LAW_TAIL = NOON + WAW + NOON;

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <!-- Deliberately NOT a bundled webfont. Pinning one makes the painted
             text use it while the measurer, which runs before the webfont has
             loaded, measures a fallback face: a 64px divergence on a 194px line
             that has nothing to do with cursive joining. That font-loading race
             is a real and separate defect; this fixture stays on locally
             available faces so it measures the joining repair alone. -->
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="28"/>
        <w:szCs w:val="28"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`;

type RunSpec = { text: string; bold?: boolean; color?: string; rtl?: boolean };

const run = ({ text, bold, color, rtl }: RunSpec): string => {
  const properties = [
    rtl ? "<w:rtl/>" : "",
    bold ? "<w:b/><w:bCs/>" : "",
    color ? `<w:color w:val="${color}"/>` : "",
  ].join("");
  return `<w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;
};

const paragraph = (runs: RunSpec[], rtl: boolean): string =>
  `<w:p><w:pPr>${rtl ? '<w:bidi/><w:jc w:val="right"/>' : ""}</w:pPr>${runs
    .map((spec) => run({ ...spec, rtl: rtl || spec.rtl }))
    .join("")}</w:p>`;

/** Paragraph 1: bold mid-word. Every split here must be repaired. */
const BOLD_SPLIT_PARAGRAPH = paragraph(
  [
    { text: OFFICE_HEAD },
    { text: OFFICE_TAIL, bold: true },
    { text: " " },
    { text: COURT_HEAD },
    { text: COURT_TAIL, bold: true },
    { text: " " },
    { text: CONSTITUTIONAL_HEAD },
    { text: CONSTITUTIONAL_TAIL, bold: true },
    { text: " " },
    { text: LAW_HEAD },
    { text: LAW_TAIL, bold: true },
  ],
  true,
);

/** Paragraph 2: colour only. Same face, so no joiner belongs anywhere here. */
const COLOUR_SPLIT_PARAGRAPH = paragraph(
  [
    { text: OFFICE_HEAD },
    { text: OFFICE_TAIL, color: "C00000" },
    { text: " " },
    { text: COURT_HEAD },
    { text: COURT_TAIL, color: "C00000" },
  ],
  true,
);

/** Paragraph 3: the Latin control, split the same way. */
const LATIN_SPLIT_PARAGRAPH = paragraph(
  [
    { text: "smlo" },
    { text: "uva", bold: true },
    { text: " o d" },
    { text: "ilo", color: "C00000" },
    { text: " 2026" },
  ],
  false,
);

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${BOLD_SPLIT_PARAGRAPH}
    ${COLOUR_SPLIT_PARAGRAPH}
    ${LATIN_SPLIT_PARAGRAPH}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const addXml = (zip: JSZip, filePath: string, contents: string): void => {
  zip.file(filePath, contents, { date: FIXED_DATE, createFolders: false });
};

const build = async (): Promise<void> => {
  const zip = new JSZip();
  addXml(zip, "[Content_Types].xml", CONTENT_TYPES);
  addXml(zip, "_rels/.rels", PACKAGE_RELS);
  addXml(zip, "word/_rels/document.xml.rels", DOCUMENT_RELS);
  addXml(zip, "word/document.xml", DOCUMENT_XML);
  addXml(zip, "word/styles.xml", STYLES_XML);

  const fixture = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
  await mkdir(OUTPUT_DIR, { recursive: true });
  await Bun.write(OUTPUT_PATH, fixture);
  console.log(`Wrote ${OUTPUT_PATH} (${fixture.byteLength} bytes)`);
};

await build();
