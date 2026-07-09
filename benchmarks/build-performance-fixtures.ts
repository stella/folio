/**
 * Deterministic DOCX fixtures for end-to-end performance scenarios.
 *
 * Generated files are committed under `tests/visual/fixtures` so Vite's dev and
 * preview servers exercise identical bytes. Content is repository-authored;
 * the embedded Arimo face comes from the Apache-2.0 `@fontsource/arimo` package.
 *
 * Run from the repository root:
 *
 *   bun benchmarks/build-performance-fixtures.ts
 */
import { readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

import { createDocx, createEmptyDocument } from "@stll/folio-core";
import type { Document } from "@stll/folio-core";
import JSZip from "jszip";

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, "..");
const FIXTURES_DIR = resolve(REPO_ROOT, "tests/visual/fixtures");
const EMBEDDED_FONT_PATH = resolve(
  REPO_ROOT,
  "packages/react/node_modules/@fontsource/arimo/files/arimo-hebrew-400-normal.woff2",
);
const PARAGRAPH_COUNT = 1500;
const FIXED_ZIP_DATE = new Date(2000, 0, 1, 0, 0, 0);
const FIXED_DOCUMENT_TIMESTAMP = "2000-01-01T00:00:00.000Z";

type BodyBlock = Document["package"]["document"]["content"][number];
type Paragraph = Extract<BodyBlock, { type: "paragraph" }>;
type Table = Extract<BodyBlock, { type: "table" }>;
type TableRow = Table["rows"][number];

const paragraph = (text: string, fontFamily = "Arial"): Paragraph => ({
  type: "paragraph",
  content: [
    {
      type: "run",
      content: [{ type: "text", text }],
      formatting: {
        fontFamily: { ascii: fontFamily, hAnsi: fontFamily },
        fontSize: 22,
      },
    },
  ],
  formatting: { lineSpacing: 276 },
});

const buildParagraphDocument = (): Document => {
  const document = createEmptyDocument();
  const content: BodyBlock[] = [];
  for (let index = 0; index < PARAGRAPH_COUNT; index += 1) {
    content.push(
      paragraph(
        `Performance paragraph ${index + 1}: repository-authored text exercises DOCX parsing, conversion, measurement, pagination, and painting.`,
      ),
    );
  }
  document.package.document.content = content;
  return document;
};

const tableCell = (content: Paragraph[]): TableRow["cells"][number] => ({
  type: "tableCell",
  content,
});

const buildLongTableDocument = (): Document => {
  const document = createEmptyDocument();
  const rows: TableRow[] = [];

  rows.push({
    type: "tableRow",
    cells: [
      tableCell(
        Array.from({ length: 80 }, (_, index) =>
          paragraph(`Split-row paragraph ${index + 1}: this single table row exceeds one page.`),
        ),
      ),
      tableCell(
        Array.from({ length: 80 }, (_, index) =>
          paragraph(`Parallel cell paragraph ${index + 1}: row fragments must remain aligned.`),
        ),
      ),
    ],
  });

  for (let index = 0; index < 120; index += 1) {
    rows.push({
      type: "tableRow",
      cells: [
        tableCell([paragraph(`Long table row ${index + 1}, first column`)]),
        tableCell([paragraph(`Long table row ${index + 1}, second column`)]),
      ],
    });
  }

  document.package.document.content = [
    paragraph("Long table with a deliberately oversized first row"),
    { type: "table", rows },
  ];
  return document;
};

const buildMixedScriptDocument = (): Document => {
  const document = createEmptyDocument();
  const content: BodyBlock[] = [];
  const lines = [
    "English and Latin: Performance baseline for international documents.",
    "العربية: هذا مستند اصطناعي لقياس أداء التخطيط والتحرير.",
    "עברית: זהו מסמך סינתטי למדידת ביצועי פריסה ועריכה.",
    "中文與日本語: 這是一個用於效能測量的合成文件。日本語の文章も含みます。",
    "Mixed: Contract סעיף 12 يتضمن شروطًا متعددة，並包含 multilingual text.",
  ] as const;

  for (let index = 0; index < 250; index += 1) {
    for (const line of lines) {
      content.push(paragraph(`${index + 1}. ${line}`, "Arimo Embedded"));
    }
  }
  document.package.document.content = content;
  return document;
};

const addEmbeddedFont = async (buffer: ArrayBuffer): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(buffer);
  const contentTypesFile = zip.file("[Content_Types].xml");
  const documentRelsFile = zip.file("word/_rels/document.xml.rels");
  if (!contentTypesFile || !documentRelsFile) {
    throw new Error("Generated DOCX is missing content types or document relationships");
  }

  const fontTable = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:font w:name="Arimo Embedded"><w:family w:val="swiss"/><w:embedRegular r:id="rIdPerfFont"/></w:font>
</w:fonts>`;
  zip.file("word/fontTable.xml", fontTable);

  const documentRels = await documentRelsFile.async("string");
  zip.file(
    "word/_rels/document.xml.rels",
    documentRels.replace(
      "</Relationships>",
      '<Relationship Id="rIdPerfFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/></Relationships>',
    ),
  );

  const fontRelationships = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPerfFont" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/arimo-hebrew.woff2"/>
</Relationships>`;
  zip.file("word/_rels/fontTable.xml.rels", fontRelationships);
  zip.file("word/fonts/arimo-hebrew.woff2", await readFile(EMBEDDED_FONT_PATH));

  const contentTypes = await contentTypesFile.async("string");
  zip.file(
    "[Content_Types].xml",
    contentTypes.replace(
      "</Types>",
      '<Default Extension="woff2" ContentType="font/woff2"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>',
    ),
  );
  return generateDeterministicDocx(zip);
};

const generateDeterministicDocx = async (zip: JSZip): Promise<ArrayBuffer> => {
  const coreProperties = zip.file("docProps/core.xml");
  if (coreProperties) {
    const xml = await coreProperties.async("string");
    zip.file(
      "docProps/core.xml",
      xml.replaceAll(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/gu, FIXED_DOCUMENT_TIMESTAMP),
    );
  }
  for (const entry of Object.values(zip.files)) {
    entry.date = FIXED_ZIP_DATE;
  }
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
};

const writeDocument = async (name: string, document: Document): Promise<void> => {
  const buffer = await createDocx(document);
  const normalized = await generateDeterministicDocx(await JSZip.loadAsync(buffer));
  await writeFile(resolve(FIXTURES_DIR, name), new Uint8Array(normalized));
};

await Promise.all([
  writeDocument("performance-1500-paragraphs.docx", buildParagraphDocument()),
  writeDocument("performance-long-split-table.docx", buildLongTableDocument()),
  (async () => {
    const buffer = await createDocx(buildMixedScriptDocument());
    const withFont = await addEmbeddedFont(buffer);
    await writeFile(
      resolve(FIXTURES_DIR, "performance-mixed-script-embedded-font.docx"),
      new Uint8Array(withFont),
    );
  })(),
]);
