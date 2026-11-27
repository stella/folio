/**
 * A batch of paragraph-structure operations (merges with and without the
 * merged paragraph's properties, splits, deletions, property sets) and
 * tables put before or after its paragraphs, over a small numbered document,
 * in every mode and after up to two earlier calls that left changes pending,
 * leaves what its operations leave one at a time, each its own call, in the
 * order the batch runs them: from the last paragraph back, with tables in a
 * gap ahead of the paragraph removal that would join the tables around that
 * gap. Tables in one gap go in one call, in batch order. Every operation the batch
 * applies applies alone too, every one it refuses for the document's state
 * is refused alone for the same reason, and the two save the same package,
 * up to revision and added paragraph ids and dates; with tables and changes
 * pending, once accepted and once rejected.
 *
 * A case opens and saves up to a dozen packages, a few hundred milliseconds
 * at most; the budget is 60 s for the stated runs, scaled by the nightly
 * factor.
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
  | { type: "properties"; block: number; properties: Properties }
  | { type: "table"; block: number; side: "before" | "after" };

type Case = {
  mode: FolioDocumentOperationMode;
  /** Operations of earlier calls, each its own, which the batch finds pending. */
  prior: readonly Step[];
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
    tableArbitrary(block),
  );

/** A table put before or after paragraph `block`, the closing one included. */
const tableArbitrary = (block: number): fc.Arbitrary<Step> =>
  fc.record({
    type: fc.constant("table" as const),
    block: fc.constant(block),
    side: fc.constantFrom<("before" | "after")[]>("before", "after"),
  });

/**
 * Two or three operations on distinct paragraphs and up to two tables put
 * beside any paragraph, in any input order.
 */
const caseArbitrary: fc.Arbitrary<Case> = fc.record({
  mode: fc.constantFrom("direct", "tracked-changes", "suggested"),
  prior: fc.array(fc.constantFrom(...TARGETS).chain(stepArbitrary), {
    maxLength: 2,
  }),
  steps: fc
    .tuple(
      fc
        .shuffledSubarray(TARGETS, { minLength: 2, maxLength: 3 })
        .chain((blocks) => fc.tuple(...blocks.map(stepArbitrary))),
      fc.array(fc.nat({ max: IDS.length - 1 }).chain(tableArbitrary), {
        maxLength: 2,
      }),
    )
    .chain(([paragraphs, tables]) => {
      const all = [...paragraphs, ...tables];
      return fc.shuffledSubarray(all, {
        minLength: all.length,
        maxLength: all.length,
      });
    }),
});

const operationOf = (
  step: Step,
  index: number | `prior-${number}`,
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
    case "table":
      return {
        id,
        type: "insertTable",
        blockId,
        position: step.side,
        rows: [[`Cell ${String(index)}`, "Value"]],
      };
  }
};

/**
 * The saved body, revision ids and the ids of paragraphs the operations add
 * renumbered in order of appearance, no dates: both are minted in the order
 * operations run.
 */
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
  const added = new Map<string, number>();
  // A join's property change names its revision by id too.
  return xml
    .replaceAll(/\sw:date="[^"]*"/gu, "")
    .replaceAll(
      /\b(w:id="|resolvesWith="\{&quot;version&quot;:\d+,&quot;value&quot;:)(\d+)/gu,
      (_, prefix: string, id: string) =>
        `${prefix}${renumbered.get(id) ?? "?"}`,
    )
    .replaceAll(/\bw14:paraId="(0000[0-9A-F]{4})"/gu, (_, id: string) => {
      if (!added.has(id)) added.set(id, added.size);
      return `w14:paraId="added-${String(added.get(id))}"`;
    });
};

type RunOptions = {
  mode: FolioDocumentOperationMode;
  batches: readonly (readonly FolioDocumentOperation[])[];
  /** Whether `body` holds the saved bodies once accepted and once rejected. */
  resolved?: boolean;
};

const reopened = async (reviewer: FolioDocxReviewer) =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
    author: "Reviewer",
    session: "canonical",
  });

/**
 * Apply each batch as its own call, each stamped a second after the last,
 * as separate calls are; what was applied and skipped, and the saved body.
 */
const run = async ({ mode, batches, resolved = false }: RunOptions) => {
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
  const blocks = reviewer.getContent().map(({ id, text }) => ({ id, text }));
  if (!resolved)
    return { applied, skipped, blocks, body: await savedBody(reviewer) };
  const accepted = await reopened(reviewer);
  accepted.acceptAll();
  const rejected = await reopened(reviewer);
  rejected.rejectAll();
  const bodies = [await savedBody(accepted), await savedBody(rejected)];
  return { applied, skipped, blocks, body: bodies.join("\n") };
};

/** Reasons only a batch has: two of its operations claim one place. */
const BATCH_ONLY_REASONS = new Set(["overlappingOperation"]);

type RankOptions = {
  /** The blocks as the batch finds them. */
  blocks: readonly string[];
  /** The gap each table lands in alone, its number of `blocks` before it. */
  gaps: ReadonlyMap<Step, number>;
};

/**
 * Where an operation runs one at a time, from the last block back, as the
 * batch finds the blocks: a paragraph's own operation at its place; the
 * tables in one gap together, just after that gap's position so they run
 * before a paragraph removal at the same position. The position is the gap
 * the table actually takes alone; an empty pending deletion remains a block
 * until accepted and does not collapse the reference order early.
 */
const rankOf = (step: Step, { blocks, gaps }: RankOptions): number => {
  const placeOf = (block: number) => blocks.indexOf(IDS[block] ?? "");
  const gap = gaps.get(step);
  if (gap === undefined) return placeOf(step.block);
  return gap + 0.25;
};

/**
 * The batch's operations alone, from the last paragraph back, the tables put
 * in one gap together in batch order: all but those it refused for another
 * of its operations. Each it applied applies alone, each it refused is
 * refused alone for the same reason.
 */
const checkCase = async ({ mode, prior, steps }: Case) => {
  const operations = steps.map(operationOf);
  const before = prior.map((step, index) => [
    operationOf(step, `prior-${index}`),
  ]);
  const found = (await run({ mode, batches: before })).blocks;
  const blocks = found.map(({ id }) => id);
  // The gap a table lands in alone. The fixture's paragraphs all hold words:
  // an empty one is one a pending deletion takes whole, and its two sides are
  // one gap, which the batch's blocks share in batch order.
  const gaps = new Map<Step, number>();
  for (const [index, step] of steps.entries()) {
    if (step.type !== "table") continue;
    const alone = await run({
      mode,
      batches: [...before, [operationOf(step, index)]],
    });
    const landed = alone.blocks.findIndex(({ id }) => !blocks.includes(id));
    const gap =
      landed < 0
        ? blocks.length
        : alone.blocks.slice(0, landed).filter(({ id }) => blocks.includes(id))
            .length;
    gaps.set(step, gap);
  }
  // Pending, the two sides of a paragraph accepting removes whole are one
  // place: with tables, the outcomes compare once accepted and once rejected.
  const resolved = mode !== "direct" && gaps.size > 0;
  const batch = await run({
    mode,
    batches: [...before, operations],
    resolved,
  });
  const batchOnly = new Set(
    batch.skipped.flatMap((skip) => {
      const [id, reason] = skip.split(": ");
      return id !== undefined && BATCH_ONLY_REASONS.has(reason ?? "")
        ? [id]
        : [];
    }),
  );
  const ranks = {
    blocks,
    gaps,
  };
  // Tracked, deleting a paragraph takes the words a pending join runs into
  // it, so the batch edits the paragraph whose mark the earlier merge or
  // deletion took before that deletion.
  const joinedInto = (step: Step) =>
    mode === "tracked-changes" &&
    step.type !== "table" &&
    prior.some(
      (earlier) =>
        (earlier.type === "merge" || earlier.type === "delete") &&
        earlier.block === step.block,
    ) &&
    steps.some(
      (other) => other.type === "delete" && other.block === step.block + 1,
    );
  const ranked = steps
    .map((step, index) => ({
      operation: operations[index] ?? panic("Expected an operation"),
      table: step.type === "table",
      // Right before that deletion, after the tables put after it.
      rank: joinedInto(step)
        ? rankOf({ ...step, block: step.block + 1 }, ranks) + 0.375
        : rankOf(step, ranks),
    }))
    .filter(({ operation }) => !batchOnly.has(operation.id))
    .toSorted((left, right) => right.rank - left.rank);
  const order: FolioDocumentOperation[][] = [];
  let tableRank: number | undefined;
  for (const { operation, rank, table } of ranked) {
    const group = order.at(-1);
    if (table && rank === tableRank && group !== undefined)
      group.push(operation);
    else order.push([operation]);
    tableRank = table ? rank : undefined;
  }
  const oneAtATime = await run({
    mode,
    batches: [...before, ...order],
    resolved,
  });
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

const table = (block: number, side: "before" | "after"): Step => ({
  type: "table",
  block,
  side,
});

/**
 * Tables beside a paragraph the batch deletes: after it and after the one
 * before it, which lands once the deletion has run; before and after it;
 * after it and the one before it once earlier calls split it and deleted
 * that one (#1356 flow 32ab93066d473a27). Beside a merge carrying properties
 * after earlier merges carrying them, the second refused (#1356 flow
 * 9d2ba4fb6c7a7388). Before the paragraph after the one a merge joins, which
 * the batch deletes: the merge then meets the table and is refused.
 */
const tableCases = (mode: FolioDocumentOperationMode): [Case][] => [
  [
    {
      mode,
      prior: [],
      steps: [merge(0), table(2, "before"), { type: "delete", block: 1 }],
    },
  ],
  [
    {
      mode,
      prior: [],
      steps: [
        table(1, "after"),
        table(2, "after"),
        { type: "delete", block: 2 },
      ],
    },
  ],
  [
    {
      mode,
      prior: [],
      steps: [
        table(2, "before"),
        table(2, "after"),
        { type: "delete", block: 2 },
      ],
    },
  ],
  [
    {
      mode,
      prior: [
        { type: "split", block: 3, first: undefined, second: undefined },
        { type: "delete", block: 2 },
      ],
      steps: [
        table(2, "after"),
        { type: "delete", block: 3 },
        table(3, "after"),
      ],
    },
  ],
  [
    {
      mode,
      prior: [
        merge(1, {
          numbering: paragraphNumberingReference({ numId: 7, ilvl: 0 }),
        }),
        merge(0, { alignment: "center" }),
      ],
      steps: [
        merge(1, { alignment: "center" }),
        table(0, "after"),
        { type: "delete", block: 2 },
      ],
    },
  ],
];

/**
 * Merges in a chain: each carrying the merged paragraph's properties; plain,
 * where the earlier join's change would go back to the paragraph's own
 * properties; and over a paragraph an earlier call deleted, carrying the
 * properties the later merge already set. Tables beside deletions
 * (`tableCases`).
 */
const MODES = ["direct", "tracked-changes"] as const;

const examples: [Case][] = MODES.flatMap(tableCases).concat(
  MODES.flatMap((mode): [Case][] => [
    [
      {
        mode,
        prior: [],
        steps: [
          merge(0, LEVEL_ONLY),
          merge(1, {
            numbering: paragraphNumberingReference({ numId: 7, ilvl: 0 }),
          }),
        ],
      },
    ],
    [{ mode, prior: [], steps: [merge(1), merge(2)] }],
    [
      {
        mode,
        prior: [{ type: "delete", block: 2 }],
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
  ]),
);

/**
 * Not yet a case: a table beside a pending join whose paragraph the batch
 * splits. The batch places it by the join as it read it; one at a time, the
 * split moves the join's first paragraph first.
 */
const placeable = ({ mode, prior, steps }: Case): boolean =>
  mode === "direct" ||
  !steps.some(({ type }) => type === "table") ||
  !steps.some(
    (step) =>
      step.type === "split" &&
      prior.some(
        (earlier) =>
          (earlier.type === "merge" || earlier.type === "delete") &&
          earlier.block === step.block,
      ),
  );

describe("a batch of paragraph-structure operations", () => {
  test(
    "leaves what its applied operations leave one at a time",
    async () => {
      await assertProperty(
        fc.asyncProperty(caseArbitrary.filter(placeable), checkCase),
        {
          id: "a batch of paragraph-structure operations equals its applied operations one at a time",
          numRuns: 60,
          examples,
        },
      );
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
