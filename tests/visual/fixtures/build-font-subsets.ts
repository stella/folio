#!/usr/bin/env bun
/**
 * Build the synthetic DOCX whose text needs several `unicode-range` subsets of
 * folio's bundled faces.
 *
 * fontsource ships each bundled face as subsets (Carlito: latin, latin-ext,
 * greek, greek-ext, cyrillic, cyrillic-ext, vietnamese; Lato: latin,
 * latin-ext) that the browser fetches only for text needing them. The
 * paragraphs name Carlito and Lato directly, so the bundled face is used on
 * every machine whether or not Calibri or Aptos is installed, and each script
 * sits in its own left-aligned paragraph long enough to wrap, plus one
 * mixed-script paragraph. No paragraph uses a Vietnamese letter, so that
 * subset stays unloaded until a test types one.
 *
 * All text is synthetic. Run: bun tests/visual/fixtures/build-font-subsets.ts
 */
import { mkdir } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

const OUTPUT_DIR = import.meta.dir;
const OUTPUT_PATH = path.join(OUTPUT_DIR, "font-subsets.docx");
/** Fixed so the fixture is byte-reproducible across rebuilds. */
const FIXED_DATE = new Date("2026-01-01T00:00:00.000Z");

const ENGLISH =
  "The Supplier shall deliver the goods to the premises of the Buyer within ten business days of the confirmation of the order.";
const CZECH =
  "Dodavatel se zavazuje dodat zboží do provozovny kupujícího do deseti pracovních dnů ode dne potvrzení objednávky; nebezpečí škody přechází okamžikem dodání.";
const POLISH =
  "Dostawca zobowiązuje się dostarczyć towar do siedziby Kupującego w ciągu dziesięciu dni roboczych od potwierdzenia zamówienia; ryzyko przechodzi z chwilą dostawy.";
const GREEK =
  "Ο Προμηθευτής υποχρεούται να παραδώσει τα αγαθά στις εγκαταστάσεις του Αγοραστή εντός δέκα εργάσιμων ημερών από την επιβεβαίωση της παραγγελίας.";
const CYRILLIC =
  "Поставщик обязуется доставить товары на объект Покупателя в течение десяти рабочих дней с момента подтверждения заказа; риск переходит в момент доставки.";
const MIXED =
  "Smluvní strany (Αγοραστής, Покупатель, Kupujący) se dohodly, že dodání zboží proběhne v souladu s článkem 4 této smlouvy a přílohou č. 2.";

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
        <w:rFonts w:ascii="Carlito" w:hAnsi="Carlito" w:cs="Carlito" w:eastAsia="Carlito"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:after="160" w:line="240" w:lineRule="auto"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`;

const escapeXml = (text: string): string =>
  text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

type ParagraphSpec = { text: string; font?: string; bold?: boolean };

const paragraph = ({ text, font, bold }: ParagraphSpec): string => {
  const properties = [
    font ? `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>` : "",
    bold ? "<w:b/><w:bCs/>" : "",
  ].join("");
  return `<w:p><w:r>${properties ? `<w:rPr>${properties}</w:rPr>` : ""}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
};

const PARAGRAPHS: ParagraphSpec[] = [
  { text: ENGLISH },
  { text: CZECH },
  { text: POLISH },
  { text: GREEK },
  { text: CYRILLIC },
  { text: MIXED },
  { text: GREEK, bold: true },
  { text: CZECH, font: "Lato" },
  { text: POLISH, font: "Lato" },
];

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${PARAGRAPHS.map(paragraph).join("\n    ")}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
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
