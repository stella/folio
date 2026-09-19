/**
 * The editor round trip over packages built here, not over corpus files: the
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
import { runEditorRoundTripInvariant } from "./lib/corpus-invariants/editor-round-trip";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const WORDPROCESSING = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DOCUMENT_PART = "word/document.xml";

const documentXml = (body: string): string =>
  `${XML_DECL}<w:document xmlns:w="${WORDPROCESSING}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${OFFICE_RELATIONSHIP}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;

const PLAIN_BODY = `<w:p w14:paraId="60000001"><w:r><w:t>Hello world</w:t></w:r></w:p>`;

/** A heading, direct run formatting and a one-cell table: three different save paths. */
const RICH_BODY = `<w:p w14:paraId="60000002"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p><w:p w14:paraId="60000003"><w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Bold italic</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr><w:p w14:paraId="60000004"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

const STYLES_XML = `${XML_DECL}<w:styles xmlns:w="${WORDPROCESSING}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style></w:styles>`;

const CORE_PROPS_XML = `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>round trip</dc:title><dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00.000Z</dcterms:modified></cp:coreProperties>`;

const PACKAGE_RELS = `${XML_DECL}<Relationships xmlns="${RELATIONSHIP_NAMESPACE}"><Relationship Id="rId1" Type="${OFFICE_RELATIONSHIP}/officeDocument" Target="${DOCUMENT_PART}"/><Relationship Id="rId2" Type="${RELATIONSHIP_NAMESPACE}/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`;

const CONTENT_TYPES = `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;

const buildPackage = (body: string): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", PACKAGE_RELS);
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECL}<Relationships xmlns="${RELATIONSHIP_NAMESPACE}"/>`,
  );
  zip.file(DOCUMENT_PART, documentXml(body));
  zip.file("word/styles.xml", STYLES_XML);
  zip.file("docProps/core.xml", CORE_PROPS_XML);
  return zip.generateAsync({ type: "arraybuffer" });
};

const invariantInput = async (body: string): Promise<CorpusInvariantInput> => {
  const buffer = await buildPackage(body);
  return {
    bytes: new Uint8Array(buffer),
    buffer,
    parsed: await parseDocx(buffer, { preloadFonts: false }),
    documentPart: DOCUMENT_PART,
    budgetMs: DEFAULT_INVARIANT_BUDGET_MS,
  };
};

describe("the editor round trip preserves the package", () => {
  test("a minimal well-formed package survives parse, edit model and save", async () => {
    const outcome = await runEditorRoundTripInvariant(await invariantInput(PLAIN_BODY));

    expect(outcome.failures).toEqual([]);
    expect(Object.keys(outcome.timings)).toEqual([
      "to-prose",
      "from-prose",
      "save",
      "parse",
      "compare",
    ]);
  });

  test("a heading, direct formatting and a table survive it too", async () => {
    const outcome = await runEditorRoundTripInvariant(await invariantInput(RICH_BODY));

    expect(outcome.failures).toEqual([]);
  });
});
