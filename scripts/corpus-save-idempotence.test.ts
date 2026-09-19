/**
 * Save idempotence over packages built here, not over corpus files: the
 * invariant has to hold for every input, and a synthetic package is the only
 * kind a test can state the expected answer for.
 */

import { describe, expect, test } from "bun:test";
import { parseDocx } from "@stll/folio-core/docx/parser";
import JSZip from "jszip";

import {
  type CorpusInvariantInput,
  DEFAULT_INVARIANT_BUDGET_MS,
} from "./lib/corpus-invariants/contract";
import {
  generalizePartPath,
  runSaveIdempotenceInvariant,
} from "./lib/corpus-invariants/save-idempotence";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const WORDPROCESSING = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DOCUMENT_PART = "word/document.xml";
const HEADER_PART = "word/header1.xml";
const MACRO_PART = "word/vbaProject.bin";
const HEADER_RELATIONSHIP_ID = "rId10";
const MACRO_RELATIONSHIP_ID = "rId11";

const STYLES_XML = `${XML_DECL}<w:styles xmlns:w="${WORDPROCESSING}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

const CORE_PROPS_XML = `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>idempotence</dc:title><dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00.000Z</dcterms:modified></cp:coreProperties>`;

/** `docProps/app.xml` is normalised by every save, so a package that has one is the interesting case. */
const APP_PROPS_XML = `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application><AppVersion>16.0000</AppVersion></Properties>`;

const HEADER_XML = `${XML_DECL}<w:hdr xmlns:w="${WORDPROCESSING}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:p w14:paraId="60000010"><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>`;

const MACRO_BYTES = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0x5a, 0x5a, 0x5a, 0x5a]);

const PACKAGE_RELS = `${XML_DECL}<Relationships xmlns="${RELATIONSHIP_NAMESPACE}"><Relationship Id="rId1" Type="${OFFICE_RELATIONSHIP}/officeDocument" Target="${DOCUMENT_PART}"/><Relationship Id="rId2" Type="${RELATIONSHIP_NAMESPACE}/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${OFFICE_RELATIONSHIP}/extended-properties" Target="docProps/app.xml"/></Relationships>`;

type ExtraPart = {
  /** Path inside the package. */
  path: string;
  /** What the archive holds at that path. */
  content: string | Uint8Array;
  /** Relationship type URI hung off `word/_rels/document.xml.rels`. */
  relationshipType: string;
  /** Relationship id the document part refers to it by. */
  relationshipId: string;
  /** `<Override>` content type for the part. */
  contentType: string;
};

const HEADER: ExtraPart = {
  path: HEADER_PART,
  content: HEADER_XML,
  relationshipType: `${OFFICE_RELATIONSHIP}/header`,
  relationshipId: HEADER_RELATIONSHIP_ID,
  contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
};

/** A part folio never models: it is carried as bytes, so it exercises the binary comparison. */
const MACRO_PROJECT: ExtraPart = {
  path: MACRO_PART,
  content: MACRO_BYTES,
  relationshipType: "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
  relationshipId: MACRO_RELATIONSHIP_ID,
  contentType: "application/vnd.ms-office.vbaProject",
};

const documentXml = (headerReference: string): string =>
  `${XML_DECL}<w:document xmlns:w="${WORDPROCESSING}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${OFFICE_RELATIONSHIP}"><w:body><w:p w14:paraId="60000001"><w:r><w:t>Hello world</w:t></w:r></w:p><w:sectPr>${headerReference}<w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

const buildPackage = (extras: readonly ExtraPart[]): Promise<ArrayBuffer> => {
  const overrides = [
    `<Override PartName="/${DOCUMENT_PART}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`,
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`,
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`,
    ...extras.map(
      ({ path, contentType }) => `<Override PartName="/${path}" ContentType="${contentType}"/>`,
    ),
  ];
  const relationships = extras.map(
    ({ path, relationshipType, relationshipId }) =>
      `<Relationship Id="${relationshipId}" Type="${relationshipType}" Target="${path.replace(/^word\//u, "")}"/>`,
  );
  const headerReference = extras.includes(HEADER)
    ? `<w:headerReference w:type="default" r:id="${HEADER_RELATIONSHIP_ID}"/>`
    : "";

  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>${overrides.join("")}</Types>`,
  );
  zip.file("_rels/.rels", PACKAGE_RELS);
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECL}<Relationships xmlns="${RELATIONSHIP_NAMESPACE}">${relationships.join("")}</Relationships>`,
  );
  zip.file(DOCUMENT_PART, documentXml(headerReference));
  zip.file("word/styles.xml", STYLES_XML);
  zip.file("docProps/core.xml", CORE_PROPS_XML);
  zip.file("docProps/app.xml", APP_PROPS_XML);
  for (const { path, content } of extras) {
    zip.file(path, content);
  }
  return zip.generateAsync({ type: "arraybuffer" });
};

const invariantInput = async (extras: readonly ExtraPart[]): Promise<CorpusInvariantInput> => {
  const buffer = await buildPackage(extras);
  return {
    bytes: new Uint8Array(buffer),
    buffer,
    parsed: await parseDocx(buffer, { preloadFonts: false }),
    documentPart: DOCUMENT_PART,
    budgetMs: DEFAULT_INVARIANT_BUDGET_MS,
  };
};

describe("saving a package folio has already saved changes nothing", () => {
  test("a minimal well-formed package is a fixed point after the first save", async () => {
    const outcome = await runSaveIdempotenceInvariant(await invariantInput([]));

    expect(outcome.failures).toEqual([]);
    expect(Object.keys(outcome.timings)).toEqual([
      "first-save",
      "reparse",
      "second-save",
      "unzip",
      "compare",
    ]);
  });

  test("a package with a header part is a fixed point too", async () => {
    const outcome = await runSaveIdempotenceInvariant(await invariantInput([HEADER]));

    expect(outcome.failures).toEqual([]);
  });

  test("a part folio carries as bytes is a fixed point too", async () => {
    const outcome = await runSaveIdempotenceInvariant(await invariantInput([MACRO_PROJECT]));

    expect(outcome.failures).toEqual([]);
  });
});

describe("generalizePartPath", () => {
  test("erases the numbers that distinguish sibling parts", () => {
    expect(generalizePartPath("word/header12.xml")).toBe("word/headerN.xml");
    expect(generalizePartPath("word/media/image3.png")).toBe("word/media/imageN.png");
  });

  test("keeps a path that carries no number", () => {
    expect(generalizePartPath("docProps/app.xml")).toBe("docProps/app.xml");
  });

  test("gives two sibling parts one signature", () => {
    expect(generalizePartPath("word/header2.xml")).toBe(generalizePartPath("word/header7.xml"));
  });
});
