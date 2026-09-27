#!/usr/bin/env bun
/**
 * Build the synthetic DOCX packages that check the review views' layout.
 *
 * `markup-views.docx` holds tracked changes in justified paragraphs: an
 * insertion, a deletion, an inserted paragraph, a deleted paragraph mark and a
 * run-size change. The two companions hold the same text written without any
 * revision: `markup-views-original.docx` as it reads with every change
 * rejected, `markup-views-final.docx` with every change accepted. Original must
 * lay out exactly like the first, No Markup and Simple Markup like the second.
 *
 * Generated from hand-written OOXML so every revision is reviewable here, and
 * so the companions are materialised independently of folio's own resolver.
 *
 * Run: bun tests/visual/fixtures/build-markup-views.ts
 */
import { mkdir } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

const OUTPUT_DIR = import.meta.dir;
/** Fixed so the fixtures are byte-reproducible across rebuilds. */
const FIXED_DATE = new Date("2026-01-01T00:00:00.000Z");

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

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

// A locally available face, as in build-cursive-face-change.ts: a bundled web
// font would add its own loading race to a comparison of two layouts.
const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="left"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>
</w:styles>`;

const REVISION = (id: number) => `w:id="${id}" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"`;
const run = (text: string, rPr = ""): string =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const inserted = (id: number, text: string): string =>
  `<w:ins ${REVISION(id)}>${run(text)}</w:ins>`;
const deleted = (id: number, text: string): string =>
  `<w:del ${REVISION(id)}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
const paragraph = (content: string, pPr = ""): string =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${content}</w:p>`;
const HEADING = '<w:pStyle w:val="Heading1"/>';
const LARGER = '<w:sz w:val="28"/><w:szCs w:val="28"/>';

const A1 = "The Supplier shall deliver the goods to the premises of the Buyer ";
const A2 = "within ten business days of the confirmation of the order";
const A3 =
  ". Risk in the goods passes to the Buyer on delivery and title passes on payment in full.";
const I1 =
  "together with all manuals, certificates of conformity, packing lists and any other documents reasonably required by the Buyer for the use, resale or export of the goods, ";
const B1 = "The Buyer shall inspect the goods ";
const D1 = "promptly after delivery and in any event within five business days, ";
const B2 =
  "and shall notify the Supplier in writing of any defect it finds, stating its nature in reasonable detail.";
const NEW_PARAGRAPH =
  "The Supplier shall replace any defective goods free of charge within ten business days of the notice, and shall bear the cost of their collection and return.";
const J1 = "The Buyer may reject goods that do not conform to the order ";
const J2 =
  "and the Supplier shall credit their price in full within thirty days of the rejection, unless the parties agree otherwise in writing.";
const SIZED = "Any notice under this agreement must be given in writing. ";
const C1 = "Either party may terminate this agreement by thirty days' written notice to the other.";

const TRACKED = [
  paragraph(run("Delivery"), HEADING),
  paragraph(run(A1) + inserted(1, I1) + run(A2) + run(A3)),
  paragraph(run(B1) + deleted(2, D1) + run(B2)),
  paragraph(inserted(4, NEW_PARAGRAPH), `<w:rPr><w:ins ${REVISION(3)}/></w:rPr>`),
  paragraph(run(J1), `<w:rPr><w:del ${REVISION(5)}/></w:rPr>`),
  paragraph(run(J2)),
  paragraph(run(SIZED, `${LARGER}<w:rPrChange ${REVISION(6)}><w:rPr/></w:rPrChange>`) + run(C1)),
];

const ORIGINAL = [
  paragraph(run("Delivery"), HEADING),
  paragraph(run(A1 + A2 + A3)),
  paragraph(run(B1 + D1 + B2)),
  paragraph(run(J1)),
  paragraph(run(J2)),
  paragraph(run(SIZED) + run(C1)),
];

const FINAL = [
  paragraph(run("Delivery"), HEADING),
  paragraph(run(A1 + I1 + A2 + A3)),
  paragraph(run(B1 + B2)),
  paragraph(run(NEW_PARAGRAPH)),
  paragraph(run(J1 + J2)),
  paragraph(run(SIZED, LARGER) + run(C1)),
];

const documentXml = (paragraphs: readonly string[]): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${paragraphs.join("")}` +
  '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>' +
  "</w:body></w:document>";

const write = async (name: string, paragraphs: readonly string[]): Promise<void> => {
  const zip = new JSZip();
  const add = (filePath: string, contents: string): void => {
    zip.file(filePath, contents, { date: FIXED_DATE, createFolders: false });
  };
  add("[Content_Types].xml", CONTENT_TYPES);
  add("_rels/.rels", PACKAGE_RELS);
  add("word/_rels/document.xml.rels", DOCUMENT_RELS);
  add("word/document.xml", documentXml(paragraphs));
  add("word/styles.xml", STYLES_XML);
  const bytes = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
  const outputPath = path.join(OUTPUT_DIR, name);
  await Bun.write(outputPath, bytes);
  console.log(`Wrote ${outputPath} (${bytes.byteLength} bytes)`);
};

await mkdir(OUTPUT_DIR, { recursive: true });
await write("markup-views.docx", TRACKED);
await write("markup-views-original.docx", ORIGINAL);
await write("markup-views-final.docx", FINAL);
