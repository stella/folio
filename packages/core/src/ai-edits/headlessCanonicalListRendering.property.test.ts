/**
 * A paragraph an operation numbers has, in the canonical reviewer's live
 * model, the list rendering the saved file reopens to.
 *
 * List rendering is derived from the paragraph's numbering and the numbering
 * definitions. Each case sets, clears, changes the level of, inherits or
 * starts a new list on a paragraph through `setBlockParagraphProperties`,
 * directly or as a tracked change, on a body paragraph, one in a table cell
 * or one in a content control, then saves and reopens: every paragraph's
 * numbering and list rendering read the same in the live model and in the
 * reopened one, and so does the indentation a numbering level adds to the
 * paragraph's formatting.
 */

import { describe, expect, mock, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import * as listNormalization from "../docx/listModelNormalization";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { canonicalReviewerSave } from "./__tests__/canonicalReviewerSave";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIListNumbering } from "./types";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml";

const paragraphXml = (text: string, properties = ""): string =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const numbered = (numId: number, ilvl: number): string =>
  `<w:numPr><w:ilvl w:val="${String(ilvl)}"/><w:numId w:val="${String(numId)}"/></w:numPr>`;

const cell = (content: string): string =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${content}</w:tc>`;

/** Plain and numbered paragraphs in the body, two table cells and a content control. */
const BODY =
  paragraphXml("Body plain") +
  paragraphXml("Body numbered", numbered(5, 0)) +
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>' +
  cell(paragraphXml("Cell plain")) +
  cell(paragraphXml("Cell numbered", numbered(5, 1))) +
  "</w:tr></w:tbl>" +
  `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${paragraphXml("Control plain")}</w:sdtContent></w:sdt>` +
  paragraphXml("Closing paragraph");

const TARGETS = [
  "Body plain",
  "Body numbered",
  "Cell plain",
  "Cell numbered",
  "Control plain",
] as const;

const level = (ilvl: number, format: string, text: string): string =>
  `<w:lvl w:ilvl="${String(ilvl)}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${String(720 * (ilvl + 1))}" w:hanging="360"/></w:pPr></w:lvl>`;

const NUMBERING =
  `${HEAD}<w:numbering xmlns:w="${W}">` +
  `<w:abstractNum w:abstractNumId="5">${level(0, "decimal", "%1.")}${level(1, "lowerLetter", "%2)")}</w:abstractNum>` +
  `<w:abstractNum w:abstractNumId="6">${level(0, "bullet", "o")}${level(1, "bullet", "-")}</w:abstractNum>` +
  '<w:num w:numId="5"><w:abstractNumId w:val="5"/></w:num>' +
  '<w:num w:numId="6"><w:abstractNumId w:val="6"/></w:num>' +
  "</w:numbering>";

const fixture = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="${MAIN}.numbering+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdN" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`,
  );
  zip.set("word/numbering.xml", NUMBERING);
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${BODY}<w:sectPr/></w:body></w:document>`,
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Reviewer",
    session: "canonical",
  });

const ilvl = fc.integer({ min: 0, max: 1 });

/** Set, clear, change the level of, inherit, or start a new list. */
const numberingArbitrary: fc.Arbitrary<FolioAIListNumbering> = fc.oneof(
  fc.record({
    kind: fc.constant("reference" as const),
    numId: fc.constantFrom(5, 6),
    ilvl,
  }),
  fc.constant({ kind: "none" as const }),
  fc.record({ kind: fc.constant("levelOnly" as const), ilvl }),
  fc.constant({ kind: "inherit" as const }),
  fc.record({
    kind: fc.constant("newList" as const),
    format: fc.constantFrom("decimal" as const, "bullet" as const),
    level: ilvl,
  }),
);

type Draw = {
  target: (typeof TARGETS)[number];
  mode: "direct" | "tracked-changes";
  numbering: FolioAIListNumbering;
};

const drawArbitrary: fc.Arbitrary<Draw> = fc.record({
  target: fc.constantFrom(...TARGETS),
  mode: fc.constantFrom("direct", "tracked-changes"),
  numbering: numberingArbitrary,
});

const asJson = (value: unknown): unknown =>
  value === undefined ? null : (JSON.parse(JSON.stringify(value)) as unknown);

const paragraphRead = (paragraph: Paragraph) => ({
  numPr: asJson(paragraph.formatting?.numPr),
  numPrFromStyle: asJson(paragraph.formatting?.numPrFromStyle),
  listRendering: asJson(paragraph.listRendering),
  // The indentation a numbering level adds to the paragraph's formatting.
  indentation: asJson({
    left: paragraph.formatting?.indentLeft,
    right: paragraph.formatting?.indentRight,
    firstLine: paragraph.formatting?.indentFirstLine,
    hanging: paragraph.formatting?.hangingIndent,
  }),
});

/** Every paragraph of the body, in tables and content controls too, by its id. */
const paragraphsOf = (document: Document) => {
  const read: Record<string, ReturnType<typeof paragraphRead>> = {};
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        read[block.paraId ?? `#${String(Object.keys(read).length)}`] =
          paragraphRead(block);
      } else if (block.type === "table") {
        for (const row of block.rows)
          for (const tableCell of row.cells) visit(tableCell.content);
      } else if (block.type === "blockSdt" || block.type === "blockCustomXml")
        visit(block.content);
    }
  };
  visit(document.package.document.content);
  return read;
};

const operationOf = (draw: Draw, blockId: string): FolioDocumentOperation => ({
  id: "number",
  type: "setBlockParagraphProperties",
  blockId,
  properties: { numbering: draw.numbering },
});

/** Apply the draw; `null` when the reviewer refuses it. */
const liveAndReopened = async (buffer: ArrayBuffer, draw: Draw) => {
  const live = await open(buffer);
  const block = live
    .snapshot()
    .blocks.find((candidate) => candidate.text === draw.target);
  if (block === undefined) throw new Error(`no block reads ${draw.target}`);
  let applied = false;
  try {
    applied =
      live.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          operations: [operationOf(draw, block.id)],
          mode: draw.mode,
        },
        { undefinedReferences: "refuse" },
      ).applied.length === 1;
  } catch {
    applied = false;
  }
  if (!applied) return null;
  const reopened = await open(await canonicalReviewerSave(live));
  return {
    live: paragraphsOf(live.toDocument()),
    reopened: paragraphsOf(reopened.toDocument()),
  };
};

const law = (buffer: ArrayBuffer, tally: { applied: number }) =>
  fc.asyncProperty(drawArbitrary, async (draw) => {
    const read = await liveAndReopened(buffer, draw);
    if (read === null) return;
    tally.applied += 1;
    expect(read.live).toEqual(read.reopened);
  });

const RUNS = 40;

/** A tracked reference that numbers a table cell's paragraph. */
const CELL_REFERENCE: Draw = {
  target: "Cell plain",
  mode: "tracked-changes",
  numbering: { kind: "reference", numId: 5, ilvl: 0 },
};

describe("numbering an operation sets", () => {
  test(
    "renders the same live and after save and reopen",
    async () => {
      const buffer = await fixture();
      await assertProperty(law(buffer, { applied: 0 }), {
        id: "list rendering reads the same live and reopened",
        numRuns: RUNS,
        examples: [[CELL_REFERENCE]],
      });
      // Most draws of a fixed sample reach the law: the reviewer applies them.
      const tally = { applied: 0 };
      await checkCoverageSample(law(buffer, tally), RUNS);
      expect(tally.applied).toBeGreaterThan(RUNS / 2);
    },
    propertyTestTimeout(120_000),
  );

  test(
    "the law catches a normalization that leaves a table's paragraphs as they were",
    async () => {
      // The test-only break: the edit normalization stops at the story's
      // own paragraphs, as it did before it walked into containers.
      const original = { ...listNormalization };
      mock.module("../docx/listModelNormalization", () => ({
        ...original,
        normalizeListModelRendering: (
          document: Document,
          options?: Parameters<typeof original.normalizeListModelRendering>[1],
        ) => {
          const tables = new Map<number, BlockContent>();
          const content = document.package.document.content;
          for (const [index, block] of content.entries())
            if (block.type === "table") tables.set(index, block);
          const result = original.normalizeListModelRendering(
            document,
            options,
          );
          const normalizedContent = result.document.package.document.content;
          if (normalizedContent.length !== content.length) return result;
          const kept = normalizedContent.map(
            (block, index) => tables.get(index) ?? block,
          );
          return {
            ...result,
            document: {
              ...result.document,
              package: {
                ...result.document.package,
                document: {
                  ...result.document.package.document,
                  content: kept,
                },
              },
            },
          };
        },
      }));
      const buffer = await fixture();
      try {
        const result = await fc.check(
          law(buffer, { applied: 0 }),
          propertyConfig({
            numRuns: 1,
            examples: [[CELL_REFERENCE]],
            reporter: () => undefined,
          }),
        );
        expect(result.failed).toBe(true);
      } finally {
        mock.module("../docx/listModelNormalization", () => original);
      }
    },
    propertyTestTimeout(60_000),
  );
});
