#!/usr/bin/env bun
import { generateDocxFixture } from "../../support/validatedDocxFixture";
/**
 * Build the synthetic DOCX an IME composition test types kana into, with the
 * run's East Asian slot naming a face the document embeds, so the kana draw
 * with the same face on every machine.
 *
 * The paragraph's run names Arial for Latin text (folio's bundled Arimo) and
 * "Kana Embedded" for East Asian text. The embedded face is the kana subset
 * of Noto Sans JP from fontsource (SIL Open Font License 1.1, see
 * fonts/noto-sans-jp-LICENSE.txt), renamed for the document.
 *
 * All text is synthetic. Run: bun tests/visual/fixtures/build-cjk-composition.ts
 */
import { readFile } from "node:fs/promises";
import path from "node:path";

import { ZipArchive } from "@stll/docx-core/zip";

const OUTPUT_DIR = import.meta.dir;
const OUTPUT_PATH = path.join(OUTPUT_DIR, "cjk-composition.docx");
const FONT_PATH = path.join(
  OUTPUT_DIR,
  "fonts",
  "noto-sans-jp-kana-400-normal.woff2",
);
/** Fixed so the fixture is byte-reproducible across rebuilds. */
const FIXED_DATE = new Date("2026-01-01T00:00:00.000Z");

const TEXT =
  "The Supplier shall deliver the goods to the premises of the Buyer within ten business days.";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="woff2" ContentType="font/woff2"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
</Relationships>`;

const FONT_TABLE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="Kana Embedded"><w:charset w:val="80"/><w:family w:val="swiss"/><w:embedRegular r:id="rIdKana"/></w:font>
</w:fonts>`;

const FONT_TABLE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdKana" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/kana-embedded.woff2"/>
</Relationships>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:sz w:val="28"/>
        <w:szCs w:val="28"/>
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

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Kana Embedded"/></w:rPr><w:t xml:space="preserve">${TEXT}</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const build = async (): Promise<void> => {
  const zip = new ZipArchive();
  const add = (filePath: string, contents: string | Uint8Array): void => {
    zip.set(filePath, contents, { modified: FIXED_DATE });
  };
  add("[Content_Types].xml", CONTENT_TYPES);
  add("_rels/.rels", PACKAGE_RELS);
  add("word/_rels/document.xml.rels", DOCUMENT_RELS);
  add("word/document.xml", DOCUMENT_XML);
  add("word/styles.xml", STYLES_XML);
  add("word/fontTable.xml", FONT_TABLE);
  add("word/_rels/fontTable.xml.rels", FONT_TABLE_RELS);
  add("word/fonts/kana-embedded.woff2", new Uint8Array(await readFile(FONT_PATH)));

  const fixture = await generateDocxFixture(zip, "cjk-composition");
  await Bun.write(OUTPUT_PATH, fixture);
  console.log(`Wrote ${OUTPUT_PATH} (${fixture.byteLength} bytes)`);
};

await build();
