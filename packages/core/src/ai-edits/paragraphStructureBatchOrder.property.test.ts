/**
 * A batch of paragraph-structure operations (merges with and without the
 * merged paragraph's properties, splits, deletions, property sets) over a
 * small numbered document, in every mode and after an earlier call that left
 * a change pending, leaves what its operations leave one at a time, each its
 * own call, in the order the batch runs them: from the last paragraph back.
 * Every operation the batch applies applies alone too, every one it refuses
 * for the document's state is refused alone for the same reason, and the two
 * save the same package, up to revision ids and dates.
 *
 * A case opens and saves a few packages, tens of milliseconds; the budget is
 * 60 s for the stated runs, scaled by the nightly factor.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { paragraphNumberingReference } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { paragraphNumberingFromSlots } from "../docx";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Paragraph } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIBlockParagraphProperties } from "./types";

const TEXTS = [
  "# Terms",
  "Numbered clause one",
  "Unnumbered body text.",
  "Numbered clause two",
  "Closing words.",
] as const;
const IDS = [
  "412DDC7F",
  "3CEFFC83",
  "31617F9A",
  "5E88C024",
  "6A0C1D2E",
] as const;
const NUMBERED = new Set([1, 3]);
/** The paragraphs operations name; the closing one keeps the story's end. */
const TARGETS = [0, 1, 2, 3];

/**
 * A heading, two clauses numbered directly from instance 7 with body text
 * between, and a closing paragraph.
 */
const source = async (): Promise<ArrayBuffer> => {
  const document = fromMarkdown(TEXTS.join("\n\n"));
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 7,
        multiLevelType: "singleLevel",
        levels: [
          {
            ilvl: 0,
            start: 1,
            numFmt: "decimal",
            lvlText: "(%1)",
            suffix: "space",
            pPr: { indentLeft: 0, indentFirstLine: 0 },
          },
        ],
      },
    ],
    nums: [{ numId: 7, abstractNumId: 7 }],
  };
  for (const [index, block] of document.package.document.content.entries()) {
    const paragraph = block as Paragraph;
    paragraph.paraId = IDS[index] ?? panic("Expected a paragraph id");
    if (NUMBERED.has(index))
      paragraph.formatting = {
        ...paragraph.formatting,
        numPr: paragraphNumberingFromSlots({ numId: 7, ilvl: 0 }),
      };
  }
  return createDocx(document);
};

const PROPERTIES: readonly FolioAIBlockParagraphProperties[] = [
  { numbering: paragraphNumberingReference({ numId: 7, ilvl: 0 }) },
  { styleId: null, numbering: { kind: "levelOnly", ilvl: 1 } },
  { alignment: "center" },
  { styleId: "Heading1" },
];

type Properties = (typeof PROPERTIES)[number];

/** One operation on paragraph `block` (an index into `IDS`). */
type Step =
  | { type: "merge"; block: number; properties: Properties | undefined }
  | {
      type: "split";
      block: number;
      first: Properties | undefined;
      second: Properties | undefined;
    }
  | { type: "delete"; block: number }
  | { type: "properties"; block: number; properties: Properties };

type Case = {
  mode: FolioDocumentOperationMode;
  /** An operation of an earlier call, which the batch finds pending. */
  prior: Step | undefined;
  steps: readonly Step[];
};

const propertiesArbitrary = fc.constantFrom(...PROPERTIES);
const maybeProperties = fc.option(propertiesArbitrary, { nil: undefined });

const stepArbitrary = (block: number): fc.Arbitrary<Step> =>
  fc.oneof(
    fc.record({
      type: fc.constant("merge" as const),
      block: fc.constant(block),
      properties: maybeProperties,
    }),
    fc.record({
      type: fc.constant("split" as const),
      block: fc.constant(block),
      first: maybeProperties,
      second: maybeProperties,
    }),
    fc.record({
      type: fc.constant("delete" as const),
      block: fc.constant(block),
    }),
    fc.record({
      type: fc.constant("properties" as const),
      block: fc.constant(block),
      properties: propertiesArbitrary,
    }),
  );

/** Two or three operations on distinct paragraphs, in any input order. */
const caseArbitrary: fc.Arbitrary<Case> = fc.record({
  mode: fc.constantFrom("direct", "tracked-changes", "suggested"),
  prior: fc.option(fc.constantFrom(...TARGETS).chain(stepArbitrary), {
    nil: undefined,
  }),
  steps: fc
    .shuffledSubarray(TARGETS, { minLength: 2, maxLength: 3 })
    .chain((blocks) => fc.tuple(...blocks.map(stepArbitrary))),
});

const operationOf = (
  step: Step,
  index: number | "prior",
): FolioDocumentOperation => {
  const id = `op-${index}`;
  const blockId = IDS[step.block] ?? panic("Expected a paragraph id");
  switch (step.type) {
    case "merge":
      return {
        id,
        type: "mergeBlockWithNext",
        blockId,
        separator: " ",
        ...(step.properties && { mergedParagraphProperties: step.properties }),
      };
    case "split":
      return {
        id,
        type: "splitBlock",
        blockId,
        offset: 3,
        ...(step.first && { firstParagraphProperties: step.first }),
        ...(step.second && { secondParagraphProperties: step.second }),
      };
    case "delete":
      return { id, type: "deleteBlock", blockId };
    case "properties":
      return {
        id,
        type: "setBlockParagraphProperties",
        blockId,
        properties: step.properties,
      };
  }
};

/** The saved body, revision ids renumbered in order of appearance, no dates. */
const savedBody = async (reviewer: FolioDocxReviewer): Promise<string> => {
  const xml =
    (await (
      await ZipArchive.read(await reviewer.toBuffer())
    )
      .entry("word/document.xml")
      ?.text()) ?? panic("Expected word/document.xml");
  const renumbered = new Map<string, number>();
  for (const [, id] of xml.matchAll(/\bw:id="(\d+)"/gu))
    if (id !== undefined && !renumbered.has(id))
      renumbered.set(id, renumbered.size);
  // A join's property change names its revision by id too.
  return xml
    .replaceAll(/\sw:date="[^"]*"/gu, "")
    .replaceAll(
      /\b(w:id="|resolvesWith="\{&quot;version&quot;:\d+,&quot;value&quot;:)(\d+)/gu,
      (_, prefix: string, id: string) =>
        `${prefix}${renumbered.get(id) ?? "?"}`,
    );
};

type RunOptions = {
  mode: FolioDocumentOperationMode;
  batches: readonly (readonly FolioDocumentOperation[])[];
};

/**
 * Apply each batch as its own call, each stamped a second after the last,
 * as separate calls are; what was applied and skipped, and the saved body.
 */
const run = async ({ mode, batches }: RunOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
    author: "Reviewer",
    session: "canonical",
  });
  const applied: string[] = [];
  const skipped: string[] = [];
  for (const [index, operations] of batches.entries()) {
    const result = reviewer.applyDocumentOperations(
      { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
      {
        undefinedReferences: "refuse",
        revisionStamp: {
          date: new Date(Date.UTC(2026, 0, 2, 3, 4, index)).toISOString(),
          idSeed: 100 * (index + 1),
        },
      },
    );
    applied.push(...result.applied.map(({ id }) => id));
    skipped.push(...result.skipped.map(({ id, reason }) => `${id}: ${reason}`));
  }
  return { applied, skipped, body: await savedBody(reviewer) };
};

/** Reasons only a batch has: two of its operations claim one place. */
const BATCH_ONLY_REASONS = new Set(["overlappingOperation"]);

/**
 * The batch's operations, each alone, from the last paragraph back: all but
 * those it refused for another of its operations. Each it applied applies
 * alone, each it refused is refused alone for the same reason.
 */
const checkCase = async ({ mode, prior, steps }: Case) => {
  const operations = steps.map(operationOf);
  const before = prior === undefined ? [] : [[operationOf(prior, "prior")]];
  const batch = await run({ mode, batches: [...before, operations] });
  const batchOnly = new Set(
    batch.skipped.flatMap((skip) => {
      const [id, reason] = skip.split(": ");
      return id !== undefined && BATCH_ONLY_REASONS.has(reason ?? "")
        ? [id]
        : [];
    }),
  );
  // Tracked, deleting a paragraph takes the words a pending join runs into
  // it, so the batch edits the paragraph whose mark the earlier merge or
  // deletion took before that deletion.
  const joinedInto = (block: number) =>
    mode === "tracked-changes" &&
    (prior?.type === "merge" || prior?.type === "delete") &&
    prior.block === block &&
    steps.some((step) => step.type === "delete" && step.block === block + 1);
  const order = operations
    .map((operation, index) => {
      const block = steps[index]?.block ?? 0;
      return { operation, rank: joinedInto(block) ? block + 1.5 : block };
    })
    .filter(({ operation }) => !batchOnly.has(operation.id))
    .toSorted((left, right) => right.rank - left.rank)
    .map(({ operation }) => [operation]);
  const oneAtATime = await run({ mode, batches: [...before, ...order] });
  expect(oneAtATime.skipped.toSorted()).toEqual(
    batch.skipped
      .filter((skip) => !batchOnly.has(skip.split(": ")[0] ?? ""))
      .toSorted(),
  );
  expect(oneAtATime.body).toEqual(batch.body);
};

const merge = (block: number, properties?: Properties): Step => ({
  type: "merge",
  block,
  properties,
});

const LEVEL_ONLY = {
  styleId: null,
  numbering: { kind: "levelOnly", ilvl: 1 },
} as const satisfies Properties;

/**
 * Merges in a chain: each carrying the merged paragraph's properties; plain,
 * where the earlier join's change would go back to the paragraph's own
 * properties; and over a paragraph an earlier call deleted, carrying the
 * properties the later merge already set.
 */
const examples: [Case][] = (["direct", "tracked-changes"] as const).flatMap(
  (mode): [Case][] => [
    [
      {
        mode,
        prior: undefined,
        steps: [
          merge(0, LEVEL_ONLY),
          merge(1, {
            numbering: paragraphNumberingReference({ numId: 7, ilvl: 0 }),
          }),
        ],
      },
    ],
    [{ mode, prior: undefined, steps: [merge(1), merge(2)] }],
    [
      {
        mode,
        prior: { type: "delete", block: 2 },
        steps: [
          merge(1, LEVEL_ONLY),
          merge(3, LEVEL_ONLY),
          {
            type: "properties",
            block: 0,
            properties: {
              numbering: paragraphNumberingReference({ numId: 7, ilvl: 0 }),
            },
          },
        ],
      },
    ],
  ],
);

describe("a batch of paragraph-structure operations", () => {
  test(
    "leaves what its applied operations leave one at a time",
    async () => {
      await assertProperty(fc.asyncProperty(caseArbitrary, checkCase), {
        id: "a batch of paragraph-structure operations equals its applied operations one at a time",
        numRuns: 60,
        examples,
      });
    },
    propertyTestTimeout(60_000),
  );
});

describe("a merge whose merged paragraph a later merge of the batch joins on", () => {
  test("tracked: is refused, as it is after that merge alone", async () => {
    const operations = [
      merge(0, LEVEL_ONLY),
      merge(1, {
        numbering: paragraphNumberingReference({ numId: 7, ilvl: 0 }),
      }),
    ].map(operationOf);
    const batch = await run({
      mode: "tracked-changes",
      batches: [operations],
    });
    expect(batch.applied).toEqual(["op-1"]);
    expect(batch.skipped).toEqual(["op-0: pendingParagraphPropertyChange"]);
    const oneAtATime = await run({
      mode: "tracked-changes",
      batches: operations.toReversed().map((operation) => [operation]),
    });
    expect(oneAtATime.skipped).toEqual([
      "op-0: pendingParagraphPropertyChange",
    ]);
    expect(oneAtATime.body).toEqual(batch.body);
  });
});
