/**
 * The schema-validity invariant, over packages built here.
 *
 * The invariant is a difference, so the case that matters is a package that
 * was already invalid: `w:jc w:val="sideways"` is a value the schema's
 * enumeration rejects, folio replays the paragraph's properties verbatim, and
 * the saved part is therefore just as invalid as the one that came in. A check
 * that validated the output alone would call that folio's defect.
 */

import { describe, expect, test } from "bun:test";
import { buildBodySequenceDocx } from "@stll/folio-core/compare/__fixtures__/body-sequence";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { createEmptyDocx, repackDocx } from "@stll/folio-core/docx/rezip";
import { unzipDocx } from "@stll/folio-core/docx/unzip";
import JSZip from "jszip";

import {
  type CorpusInvariantInput,
  DEFAULT_INVARIANT_BUDGET_MS,
} from "./lib/corpus-invariants/contract";
import {
  generalizePartPath,
  runSchemaValidityInvariant,
} from "./lib/corpus-invariants/schema-validity";
import {
  loadSchemaGraph,
  SCHEMA_VIOLATION_KINDS,
  validateOoxmlPart,
} from "./lib/corpus-schema-validator";

const DOCUMENT_PART = "word/document.xml";
const HEADER_PART = "word/header1.xml";

const WORDPROCESSING = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const OFFICE_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MARKUP_COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const WORDML_2010 = "http://schemas.microsoft.com/office/word/2010/wordml";

/** A value `ST_Jc` does not enumerate, as `corpus-schema-validator.test.ts` proves. */
const REJECTED_JUSTIFICATION = "sideways";

const INVALID_BODY =
  `<w:p w14:paraId="60000001"><w:pPr><w:jc w:val="${REJECTED_JUSTIFICATION}"/></w:pPr>` +
  `<w:r><w:t xml:space="preserve">Body.</w:t></w:r></w:p>`;

/** `createEmptyDocx` supplies every part but the body, which is the only one under test. */
const packageWithBody = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    DOCUMENT_PART,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${WORDPROCESSING}" xmlns:r="${OFFICE_RELATIONSHIP}" ` +
      `xmlns:mc="${MARKUP_COMPATIBILITY}" xmlns:w14="${WORDML_2010}" mc:Ignorable="w14">` +
      `<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>` +
      `</w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const inputFor = async (buffer: ArrayBuffer): Promise<CorpusInvariantInput> => ({
  bytes: new Uint8Array(buffer),
  buffer,
  parsed: await parseDocx(buffer, { preloadFonts: false }),
  documentPart: DOCUMENT_PART,
  budgetMs: DEFAULT_INVARIANT_BUDGET_MS,
});

const partsOf = async (buffer: ArrayBuffer): Promise<ReadonlyMap<string, string>> =>
  (await unzipDocx(buffer, { extractAllXml: true })).allXml;

const graph = await loadSchemaGraph();

describe("generalizePartPath", () => {
  test("erases the ordinal that distinguishes sibling parts", () => {
    expect(generalizePartPath("word/header12.xml")).toBe("word/headerN.xml");
    expect(generalizePartPath("word/media/image3.png")).toBe("word/media/imageN.png");
    expect(generalizePartPath("word/theme/theme1.xml")).toBe("word/theme/themeN.xml");
  });

  test("keeps a path that carries no ordinal", () => {
    expect(generalizePartPath(DOCUMENT_PART)).toBe(DOCUMENT_PART);
  });
});

describe("a save may not introduce a schema violation", () => {
  test("a clean minimal package reports nothing", async () => {
    const outcome = await runSchemaValidityInvariant(
      await inputFor(await buildBodySequenceDocx([{ kind: "paragraph", text: "Body." }])),
    );

    expect(outcome.failures).toEqual([]);
    expect(Object.keys(outcome.timings)).toEqual([
      "load-schema",
      "validate-input",
      "save",
      "validate-output",
      "compare",
    ]);
  });

  test("a violation the input already carried is not folio's to answer for", async () => {
    const input = await inputFor(await packageWithBody(INVALID_BODY));
    const saved = await repackDocx(input.parsed, { updateModifiedDate: false });

    // Without this the test would pass on a save that merely dropped the
    // offending attribute, which is not the difference semantics under test.
    const savedXml = (await partsOf(saved)).get(DOCUMENT_PART) ?? "";
    const rejectedValues = validateOoxmlPart({ graph, xml: savedXml }).filter(
      ({ kind }) => kind === SCHEMA_VIOLATION_KINDS.badEnumValue,
    );
    expect(rejectedValues).toMatchObject([{ path: "document/body/p/pPr/jc" }]);

    expect((await runSchemaValidityInvariant(input)).failures).toEqual([]);
  });

  /**
   * A smoke assertion: folio replays a header verbatim and never writes one the
   * input did not have, so no synthetic package makes it lose a header's
   * validity. What this pins is that the header is written and validated at
   * all, not that a header defect would be caught.
   */
  test("a header-bearing package reports nothing (smoke)", async () => {
    const buffer = await buildBodySequenceDocx([{ kind: "paragraph", text: "Body." }], {
      header: [{ kind: "paragraph", text: "Header paragraph." }],
    });
    const input = await inputFor(buffer);

    const saved = await repackDocx(input.parsed, { updateModifiedDate: false });
    expect([...(await partsOf(saved)).keys()]).toContain(HEADER_PART);
    expect(generalizePartPath(HEADER_PART)).toBe("word/headerN.xml");

    expect((await runSchemaValidityInvariant(input)).failures).toEqual([]);
  });
});
