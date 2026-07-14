#!/usr/bin/env bun
/**
 * Build the synthetic DOCX used by the Microsoft Word line-endpoint baseline.
 *
 * The fixture is generated from hand-written OOXML so its contents and layout
 * controls remain reviewable. Keep the ZIP timestamps fixed: the captured
 * manifest is bound to the exact DOCX SHA-256.
 */
import { mkdir } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

const OUTPUT_DIR = import.meta.dir;
const OUTPUT_PATH = path.join(OUTPUT_DIR, "word-hyphenation-hanging.docx");
const FIXED_DATE = new Date("2026-01-01T00:00:00.000Z");

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`;

const SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:autoHyphenation/>
  <w:doNotHyphenateCaps/>
  <w:consecutiveHyphenLimit w:val="1"/>
  <w:noLineBreaksBefore w:lang="ja-JP" w:val="、。）］※"/>
  <w:noLineBreaksAfter w:lang="ja-JP" w:val="（［"/>
</w:settings>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Yu Gothic"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="ja-JP"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`;

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="singleLevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

const narrowParagraph = (text: string, language: string, extraProperties = ""): string => `
    <w:p>
      <w:pPr>
        <w:ind w:right="6480"/>
        ${extraProperties}
      </w:pPr>
      <w:r>
        <w:rPr><w:lang w:val="${language}" w:eastAsia="${language}"/></w:rPr>
        <w:t>${text}</w:t>
      </w:r>
    </w:p>`;

const narrowParagraphRuns = (runs: string, extraProperties = ""): string => `
    <w:p>
      <w:pPr>
        <w:ind w:right="6480"/>
        ${extraProperties}
      </w:pPr>
      ${runs}
    </w:p>`;

const hangingProbeParagraph = (text: string, extraProperties: string): string => `
    <w:p>
      <w:pPr>
        <w:ind w:right="6500"/>
        ${extraProperties}
      </w:pPr>
      <w:r>
        <w:rPr><w:lang w:val="ja-JP" w:eastAsia="ja-JP"/></w:rPr>
        <w:t>${text}</w:t>
      </w:r>
    </w:p>`;

const hangingProbeRuns = (runs: string): string => `
    <w:p>
      <w:pPr>
        <w:ind w:right="6500"/>
        <w:kinsoku/><w:overflowPunct/>
      </w:pPr>
      ${runs}
    </w:p>`;

const heading = (text: string): string => `
    <w:p>
      <w:r>
        <w:rPr><w:b/><w:lang w:val="en-US"/></w:rPr>
        <w:t>${text}</w:t>
      </w:r>
    </w:p>`;

const commonLayoutParagraph = (text: string, extraProperties: string): string => `
    <w:p>
      <w:pPr>
        <w:suppressAutoHyphens/>
        ${extraProperties}
      </w:pPr>
      <w:r><w:t>${text}</w:t></w:r>
    </w:p>`;

const tabStopProbe = (): string => `
    <w:p>
      <w:pPr>
        <w:suppressAutoHyphens/>
        <w:ind w:right="2880"/>
        <w:tabs>
          <w:tab w:val="left" w:pos="1440"/>
          <w:tab w:val="right" w:pos="5760"/>
        </w:tabs>
      </w:pPr>
      <w:r><w:t>Service item</w:t><w:tab/><w:t>Annual renewal and support plan</w:t><w:tab/><w:t>1,250.00</w:t></w:r>
    </w:p>`;

const numberedProbe = (): string => `
    <w:p>
      <w:pPr>
        <w:suppressAutoHyphens/>
        <w:ind w:right="3600"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr>
      <w:r><w:t>The supplier must preserve records, provide audit access, and notify the customer before deleting retained material.</w:t></w:r>
    </w:p>`;

const tableCellProbe = (): string => `
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="4320" w:type="dxa"/>
        <w:tblLayout w:type="fixed"/>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="4320" w:type="dxa"/></w:tcPr>
          <w:p>
            <w:pPr><w:suppressAutoHyphens/></w:pPr>
            <w:r><w:t>Payment is due within thirty days after receipt of a valid invoice and supporting records.</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>`;

const mixedFormattingProbe = (): string => `
    <w:p>
      <w:pPr><w:suppressAutoHyphens/><w:ind w:right="4320"/></w:pPr>
      <w:r><w:t xml:space="preserve">The agreement requires </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>written notice</w:t></w:r>
      <w:r><w:t xml:space="preserve"> before either party may </w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>terminate for convenience</w:t></w:r>
      <w:r><w:t>. Additional obligations continue.</w:t></w:r>
    </w:p>`;

const pageBreak = (): string => `
    <w:p><w:r><w:br w:type="page"/></w:r></w:p>`;

const cases = [
  `${heading("EN automatic")}${narrowParagraph(
    "Representation internationalization characterization demonstration compatibility.",
    "en-US",
  )}`,
  `${heading("EN paragraph suppression")}${narrowParagraph(
    "Representation internationalization characterization demonstration compatibility.",
    "en-US",
    "<w:suppressAutoHyphens/>",
  )}`,
  `${heading("EN all caps")}${narrowParagraph(
    "INTERNATIONALIZATION CHARACTERIZATION REPRESENTATION COMPATIBILITY",
    "en-US",
  )}`,
  `${heading("EN consecutive limit")}${narrowParagraph(
    "Internationalization characterization representation demonstration compatibility localization.",
    "en-US",
  )}`,
  `${heading("Czech automatic")}${narrowParagraph(
    "Nejpravděpodobnější charakteristika internacionalizace dokumentace.",
    "cs-CZ",
  )}`,
  `${heading("British English automatic")}${narrowParagraph(
    "Characterisation internationalisation localisation standardisation documentation.",
    "en-GB",
  )}`,
  `${heading("Japanese hanging punctuation")}${narrowParagraph(
    "契約条件（重要事項）。契約条件、重要事項）契約条件。",
    "ja-JP",
    "<w:kinsoku/><w:overflowPunct/>",
  )}`,
  `${heading("Custom prohibited character")}${narrowParagraph(
    "甲乙丙丁※甲乙丙丁※甲乙丙丁※甲乙丙丁",
    "ja-JP",
    "<w:kinsoku/><w:overflowPunct/>",
  )}`,
  `${heading("EN split formatting")}${narrowParagraphRuns(`
      <w:r>
        <w:rPr><w:lang w:val="en-US"/></w:rPr>
        <w:t xml:space="preserve">Representation internatio</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:color w:val="C00000"/><w:lang w:val="en-US"/></w:rPr>
        <w:t>nalization characterization demonstration compatibility.</w:t>
      </w:r>`)}`,
  `${heading("Japanese overflow disabled")}${hangingProbeParagraph(
    "甲乙丙丁戊己庚辛壬癸子丑寅。甲乙丙丁",
    '<w:kinsoku/><w:overflowPunct w:val="0"/>',
  )}`,
  `${heading("Japanese overflow enabled")}${hangingProbeParagraph(
    "甲乙丙丁戊己庚辛壬癸子丑寅。甲乙丙丁",
    "<w:kinsoku/><w:overflowPunct/>",
  )}`,
  `${heading("Japanese closing pair")}${hangingProbeParagraph(
    "甲乙丙丁戊己庚辛壬癸子丑寅。」甲乙丙丁",
    "<w:kinsoku/><w:overflowPunct/>",
  )}`,
  `${heading("Japanese split punctuation")}${hangingProbeRuns(
    `
      <w:r>
        <w:rPr><w:lang w:val="ja-JP" w:eastAsia="ja-JP"/></w:rPr>
        <w:t>甲乙丙丁戊己庚辛壬癸子丑寅</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:color w:val="C00000"/><w:lang w:val="ja-JP" w:eastAsia="ja-JP"/></w:rPr>
        <w:t>。</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:lang w:val="ja-JP" w:eastAsia="ja-JP"/></w:rPr>
        <w:t>甲乙丙丁</w:t>
      </w:r>`,
  )}`,
  `${heading("Justified common paragraph")}${commonLayoutParagraph(
    "Contract clauses allocate risk between parties and describe notice requirements, payment timing, remedies, and termination rights.",
    '<w:ind w:right="3600"/><w:jc w:val="both"/>',
  )}`,
  `${heading("First-line paragraph indent")}${commonLayoutParagraph(
    "Contract clauses allocate risk between parties and describe notice requirements, payment timing, remedies, and termination rights.",
    '<w:ind w:left="720" w:right="3600" w:firstLine="720"/>',
  )}`,
  `${heading("Explicit tab stops")}${tabStopProbe()}`,
  `${heading("Wrapped numbered item")}${numberedProbe()}`,
  `${heading("Constrained table cell")}${tableCellProbe()}`,
  `${heading("Mixed emphasis runs")}${mixedFormattingProbe()}`,
  `${heading("Left-aligned control paragraph")}${commonLayoutParagraph(
    "Contract clauses allocate risk between parties and describe notice requirements, payment timing, remedies, and termination rights.",
    '<w:ind w:right="3600"/>',
  )}`,
];

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${cases.join(pageBreak())}
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
  addXml(zip, "word/settings.xml", SETTINGS_XML);
  addXml(zip, "word/styles.xml", STYLES_XML);
  addXml(zip, "word/numbering.xml", NUMBERING_XML);

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
