/**
 * A paragraph an operation gives properties has, in the canonical reviewer's
 * live model, the formatting the saved file reopens to.
 *
 * Each case sets a paragraph's properties (spacing, indentation, alignment)
 * through `setBlockParagraphProperties`, `splitBlock` or `mergeBlockWithNext`,
 * directly or as a tracked change, on a body paragraph or one in a table
 * cell, then saves and reopens: every paragraph's formatting, and the
 * formatting each tracked property change records, reads the same in the
 * live model and in the reopened one. A case may follow the operation with a
 * direct property edit of the same paragraph, which a property change the
 * operation left pending then records as the paragraph's current properties.
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
import * as snapshots from "../../../docx-core/src/ops/paragraphPropertySnapshot";
import * as patchAsRead from "../canonical/paragraphPatchAsRead";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphFormatting,
  ParagraphPropertyChange,
} from "../types/document";
import { FolioDocxReviewer } from "./headless";
import type {
  FolioAIBlockParagraphProperties,
  FolioAIParagraphIndentation,
} from "./types";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const paragraphXml = (text: string, properties = ""): string =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** Body paragraphs with and without stated properties, and a table cell's two. */
const BODY =
  paragraphXml("First paragraph", '<w:spacing w:before="240" w:after="60"/>') +
  paragraphXml("Second paragraph", '<w:ind w:left="720" w:hanging="360"/>') +
  paragraphXml("Third paragraph", '<w:jc w:val="center"/>') +
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
  paragraphXml("Cell first", '<w:spacing w:after="120"/>') +
  paragraphXml("Cell second") +
  "</w:tc></w:tr></w:tbl>" +
  paragraphXml("Closing paragraph");

/** The paragraphs an operation targets, each followed by one it can join. */
const TARGETS = ["First paragraph", "Second paragraph", "Cell first"] as const;

const fixture = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body>${BODY}<w:sectPr/></w:body></w:document>`,
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

const twips = fc.integer({ min: 0, max: 1440 });

const propertiesArbitrary: fc.Arbitrary<FolioAIBlockParagraphProperties> = fc
  .record(
    {
      spacing: fc.option(
        fc.record(
          {
            spaceBefore: twips,
            spaceAfter: twips,
            lineSpacing: fc.integer({ min: 120, max: 720 }),
            lineSpacingRule: fc.constantFrom("auto", "exact", "atLeast"),
          },
          { requiredKeys: [] },
        ),
        { nil: null },
      ),
      indentation: fc.option(
        fc.record(
          {
            indentLeft: twips,
            indentRight: twips,
            indentFirstLine: fc.integer({ min: -720, max: 720 }),
            hangingIndent: fc.boolean(),
          },
          { requiredKeys: [] },
        ),
        { nil: null },
      ),
      alignment: fc.option(fc.constantFrom("start", "center", "end", "both"), {
        nil: null,
      }),
    },
    { requiredKeys: [] },
  )
  // An operation states at least one property; an empty set is no request.
  .filter((properties) => Object.keys(properties).length > 0);

type Draw = {
  kind: "set" | "split" | "merge";
  target: (typeof TARGETS)[number];
  mode: "direct" | "tracked-changes";
  first: FolioAIBlockParagraphProperties;
  second: FolioAIBlockParagraphProperties;
  /** A direct property edit of the same paragraph that follows, if any. */
  afterwards: FolioAIBlockParagraphProperties | null;
};

const drawArbitrary: fc.Arbitrary<Draw> = fc.record({
  kind: fc.constantFrom("set", "split", "merge"),
  target: fc.constantFrom(...TARGETS),
  mode: fc.constantFrom("direct", "tracked-changes"),
  first: propertiesArbitrary,
  second: propertiesArbitrary,
  afterwards: fc.option(propertiesArbitrary, { nil: null }),
});

const operationOf = (draw: Draw, blockId: string): FolioDocumentOperation => {
  switch (draw.kind) {
    case "set":
      return {
        id: "edit",
        type: "setBlockParagraphProperties",
        blockId,
        properties: draw.first,
      };
    case "split":
      return {
        id: "edit",
        type: "splitBlock",
        blockId,
        offset: 4,
        firstParagraphProperties: draw.first,
        secondParagraphProperties: draw.second,
      };
    case "merge":
      return {
        id: "edit",
        type: "mergeBlockWithNext",
        blockId,
        mergedParagraphProperties: draw.first,
      };
  }
};

/** A formatting record as JSON, without the source bytes it may replay. */
const formattingRead = (
  formatting: ParagraphFormatting | undefined,
): unknown => {
  if (formatting === undefined) return null;
  const modeled = Object.fromEntries(
    Object.entries(formatting).filter(([key]) => key !== "sourceXml"),
  );
  return JSON.parse(JSON.stringify(modeled)) as unknown;
};

const paragraphRead = (paragraph: Paragraph) => ({
  formatting: formattingRead(paragraph.formatting),
  propertyChanges: (paragraph.propertyChanges ?? []).map((change) => ({
    previous: formattingRead(change.previousFormatting),
    current: formattingRead(change.currentFormatting),
  })),
});

/** Every paragraph of the body, table cells included, by its id. */
const paragraphsOf = (document: Document) => {
  const read: Record<string, ReturnType<typeof paragraphRead>> = {};
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        read[block.paraId ?? `#${Object.keys(read).length}`] =
          paragraphRead(block);
      } else if (block.type === "table") {
        for (const row of block.rows)
          for (const cell of row.cells) visit(cell.content);
      } else if (block.type === "blockSdt") visit(block.content);
    }
  };
  visit(document.package.document.content);
  return read;
};

/** Apply the draw; `null` when the reviewer refuses it. */
const liveAndReopened = async (buffer: ArrayBuffer, draw: Draw) => {
  const live = await open(buffer);
  const block = live
    .snapshot()
    .blocks.find((candidate) => candidate.text === draw.target);
  if (block === undefined) throw new Error(`no block reads ${draw.target}`);
  // A refusal comes back as a skipped operation; anything thrown fails the law.
  const applies = (
    operation: FolioDocumentOperation,
    mode: Draw["mode"],
  ): boolean =>
    live.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations: [operation],
        mode,
      },
      { undefinedReferences: "refuse" },
    ).applied.length === 1;
  if (!applies(operationOf(draw, block.id), draw.mode)) return null;
  // The paragraph the operation left under the id, edited directly; whether
  // it applies or is refused, the law reads the document it leaves.
  const { afterwards } = draw;
  if (
    afterwards !== null &&
    live.snapshot().blocks.some((candidate) => candidate.id === block.id)
  )
    applies(
      {
        id: "afterwards",
        type: "setBlockParagraphProperties",
        blockId: block.id,
        properties: afterwards,
      },
      "direct",
    );
  const reopened = await open(await live.toBuffer());
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

/** A tracked merge that states the merged paragraph's spacing. */
const SPACED_MERGE: Draw = {
  kind: "merge",
  target: "First paragraph",
  mode: "tracked-changes",
  first: { spacing: { spaceBefore: 120, spaceAfter: 120 } },
  second: {},
  afterwards: null,
};

/** A direct edit of a paragraph whose tracked property change is pending. */
const DIRECT_AFTER_TRACKED: Draw = {
  kind: "set",
  target: "Second paragraph",
  mode: "tracked-changes",
  first: { alignment: "center" },
  second: {},
  afterwards: { spacing: { spaceBefore: 240 } },
};

describe("paragraph properties an operation sets", () => {
  test(
    "read the same live and after save and reopen",
    async () => {
      const buffer = await fixture();
      await assertProperty(law(buffer, { applied: 0 }), {
        id: "paragraph properties read the same live and reopened",
        numRuns: RUNS,
        examples: [[SPACED_MERGE], [DIRECT_AFTER_TRACKED]],
      });
      // Most draws of a fixed sample reach the law: the reviewer applies them.
      const tally = { applied: 0 };
      await checkCoverageSample(law(buffer, tally), RUNS);
      expect(tally.applied).toBeGreaterThan(RUNS / 2);
    },
    propertyTestTimeout(120_000),
  );

  test(
    "the law catches an edit whose formatting is left as the operation stated it",
    async () => {
      // The test-only break: an operation's patch states the paragraph as
      // the operation gave it, without the fields the parser derives from
      // what the save writes.
      const original = { ...patchAsRead };
      mock.module("../canonical/paragraphPatchAsRead", () => ({
        ...original,
        spacingSidesAsRead: () => undefined,
        indentationAsRead: (indentation: FolioAIParagraphIndentation) =>
          indentation,
      }));
      const buffer = await fixture();
      try {
        const result = await fc.check(
          law(buffer, { applied: 0 }),
          propertyConfig({
            numRuns: 1,
            examples: [[SPACED_MERGE]],
            reporter: () => undefined,
          }),
        );
        expect(result.failed).toBe(true);
      } finally {
        mock.module("../canonical/paragraphPatchAsRead", () => original);
      }
    },
    propertyTestTimeout(60_000),
  );

  test(
    "the law catches a direct edit that leaves a pending change's current properties as recorded",
    async () => {
      // The test-only break: a direct edit leaves the pending change as the
      // tracked edit recorded it.
      const original = { ...snapshots };
      mock.module(
        "../../../docx-core/src/ops/paragraphPropertySnapshot",
        () => ({
          ...original,
          changesRecordingAuthored: (
            changes: readonly ParagraphPropertyChange[],
          ) => [...changes],
        }),
      );
      const buffer = await fixture();
      try {
        const result = await fc.check(
          law(buffer, { applied: 0 }),
          propertyConfig({
            numRuns: 1,
            examples: [[DIRECT_AFTER_TRACKED]],
            reporter: () => undefined,
          }),
        );
        expect(result.failed).toBe(true);
      } finally {
        mock.module(
          "../../../docx-core/src/ops/paragraphPropertySnapshot",
          () => original,
        );
      }
    },
    propertyTestTimeout(60_000),
  );
});
