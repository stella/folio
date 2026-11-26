/**
 * The blocks one batch puts beside shared anchors land in the gaps the same
 * insertions land in made one at a time, and the blocks a batch puts in one
 * gap keep the batch's order: a later insertion after a block lands after
 * what the batch already put there. That holds tracked and directly alike,
 * and accepting the tracked batch gives what the direct one gives.
 *
 * Small documents of three or four paragraphs, some with a table, start
 * from pending changes another reviewer left: an emptied paragraph whose
 * break is pending deletion (it reads as one paragraph with the next), a
 * tracked join of two paragraphs, or the story's last paragraph deleted, so
 * the break of the one before it goes. A batch puts two or three paragraphs
 * or tables before or after a few anchors (paragraphs, pending ones
 * included, and a table cell, which puts blocks beside its table), and may
 * delete a paragraph too.
 *
 * Where each insertion lands is read from the live document as the
 * insertions run one at a time, in batch order: its gap is the number of
 * the document's blocks before it, those before a block it is put before
 * that accepting removes whole aside (both sides of those are one place).
 * The expected order merges the document's blocks with the inserted ones by
 * gap, then batch order, and keeps the blocks the accepted batch still has.
 *
 * A case opens the document about six times, some 0.3 s; the budget is
 * 120 s for the stated runs, scaled by the nightly factor.
 */

import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

type Setup = "none" | "emptiedJoin" | "trackedJoin" | "finalDeleted";
type Insertion = {
  kind: "paragraph" | "table";
  side: "before" | "after";
  /** Which of the case's anchors (`anchorsOf`) it names. */
  anchor: number;
};
type Case = {
  paragraphs: number;
  /** The paragraph a table follows, if any. */
  tableAfter: number | undefined;
  setup: Setup;
  /** The paragraph the setup changes. */
  target: number;
  insertions: Insertion[];
  /** The anchor the batch deletes too, if any. */
  deletes: number | undefined;
};

const LABEL = /\b(?:P\d|I\d|TBL)\b/gu;

const open = async ({ paragraphs, tableAfter }: Case) => {
  const parts: string[] = [];
  for (let index = 0; index < paragraphs; index++) {
    parts.push(`P${index} clause.`);
    if (index === tableAfter)
      parts.push("| TBL a | TBL b |\n| --- | --- |\n| TBL c | TBL d |");
  }
  const { docx } = await ensureParaIds(
    await createDocx(fromMarkdown(parts.join("\n\n"))),
  );
  return FolioDocxReviewer.fromBuffer(docx, { author: "Earlier" });
};

/** The case's document with the pending changes its setup leaves, reopened. */
const prepared = async (setup: Case) => {
  const earlier = await open(setup);
  const ids = earlier
    .snapshot()
    .blocks.filter(({ table }) => table === undefined)
    .map(({ id }) => id);
  const target = ids[Math.min(setup.target, ids.length - 1)] ?? "";
  // A join of a paragraph a table follows is refused, which leaves no change.
  const apply = (operation: FolioDocumentOperation) =>
    earlier.applyOperations([operation], { mode: "tracked-changes" });
  switch (setup.setup) {
    case "none":
      break;
    case "emptiedJoin":
      apply({ id: "empty", type: "deleteBlock", blockId: target });
      for (const change of earlier.getChanges())
        if (change.type !== "paragraphMarkDeleted")
          earlier.acceptChange(change);
      break;
    case "trackedJoin":
      apply({
        id: "join",
        type: "mergeBlockWithNext",
        blockId: target,
        separator: " ",
      });
      break;
    case "finalDeleted":
      apply({ id: "final", type: "deleteBlock", blockId: ids.at(-1) ?? "" });
      break;
  }
  return FolioDocxReviewer.fromBuffer(await earlier.toBuffer(), {
    author: "Reviewer",
  });
};

/** The anchors a case names: its paragraphs, then the table's first cell. */
const anchorsOf = (reviewer: FolioDocxReviewer) => {
  const blocks = reviewer.snapshot().blocks;
  const cell = blocks.find(({ table }) => table !== undefined);
  return [
    ...blocks.filter(({ table }) => table === undefined).map(({ id }) => id),
    ...(cell === undefined ? [] : [cell.id]),
  ];
};

const operationOf = (
  { kind, side, anchor }: Insertion,
  index: number,
  anchors: readonly string[],
): FolioDocumentOperation => {
  const blockId = anchors[anchor % anchors.length] ?? "";
  const id = `I${index}`;
  return kind === "table"
    ? {
        id,
        type: "insertTable",
        blockId,
        position: side,
        rows: [[`${id} term`]],
      }
    : {
        id,
        type: side === "after" ? "insertAfterBlock" : "insertBeforeBlock",
        blockId,
        text: `${id} alone.`,
      };
};

/** Each block's labels in order, a table's cells read once. */
const labelsOf = (reviewer: FolioDocxReviewer): string[] => {
  const labels: string[] = [];
  for (const { text } of reviewer.snapshot().blocks)
    for (const label of text.match(LABEL) ?? [])
      if (label !== labels.at(-1)) labels.push(label);
  return labels;
};

/** Apply `operations` to the case's prepared document in `mode`. */
const run = async (
  setup: Case,
  mode: "direct" | "tracked-changes",
  operations: (anchors: readonly string[]) => FolioDocumentOperation[],
) => {
  const reviewer = await prepared(setup);
  const before = reviewer
    .snapshot()
    .blocks.map(({ id, text }) => ({ id, label: text.match(LABEL)?.[0] }));
  const markDeleted = new Set(
    reviewer
      .getChanges()
      .flatMap((change) =>
        change.type === "paragraphMarkDeleted" && change.blockId !== undefined
          ? [change.blockId]
          : [],
      ),
  );
  const anchors = anchorsOf(reviewer);
  const result = reviewer.applyOperations(operations(anchors), { mode });
  return {
    reviewer,
    before,
    anchors,
    markDeleted,
    skipped: result.skipped.map(({ id }) => id),
  };
};

/**
 * The gap each insertion lands in made one at a time, in batch order: how
 * many of the document's blocks stand before it, the blocks earlier ones
 * put there aside. Put before a block, it is in the gap before the blocks
 * right before it that accepting removes whole. Undefined for one refused
 * at its turn.
 */
const gapsOneAtATime = async (setup: Case) => {
  const accepted = await prepared(setup);
  accepted.acceptAll();
  const kept = new Set(accepted.snapshot().blocks.map(({ id }) => id));
  const reviewer = await prepared(setup);
  const anchors = anchorsOf(reviewer);
  const original = reviewer.snapshot().blocks;
  const ids = new Set(original.map(({ id }) => id));
  const goesWhole = new Set(
    original.flatMap(({ id, text }) =>
      text === "" && !kept.has(id) ? [id] : [],
    ),
  );
  return setup.insertions.map((insertion, at) => {
    const { skipped } = reviewer.applyOperations(
      [operationOf(insertion, at, anchors)],
      { mode: "tracked-changes" },
    );
    if (skipped.length > 0) return undefined;
    const blocks = reviewer.snapshot().blocks;
    let end = blocks.findIndex(({ text }) => text.includes(`I${at} `));
    if (insertion.side === "before")
      while (end > 0 && goesWhole.has(blocks[end - 1]?.id ?? "")) end -= 1;
    return blocks.slice(0, end).filter(({ id }) => ids.has(id)).length;
  });
};

/** The batch's blocks once accepted, and the order the gaps they land in give. */
const replay = async (setup: Case) => {
  const gaps = await gapsOneAtATime(setup);
  // An insertion refused at its turn, a table beside a pending break say, is
  // left out of the batch.
  const kept = setup.insertions.flatMap((insertion, at) => {
    const gap = gaps[at];
    return gap === undefined ? [] : [{ insertion, at, gap }];
  });
  const operations = (anchors: readonly string[]) => [
    ...kept.map(({ insertion, at }) => operationOf(insertion, at, anchors)),
    ...(setup.deletes === undefined
      ? []
      : [
          {
            id: "delete",
            type: "deleteBlock",
            blockId: anchors[setup.deletes % anchors.length] ?? "",
          } satisfies FolioDocumentOperation,
        ]),
  ];
  const tracked = await run(setup, "tracked-changes", operations);
  const direct = await run(setup, "direct", operations);
  const deletedId =
    setup.deletes === undefined
      ? undefined
      : tracked.anchors[setup.deletes % tracked.anchors.length];
  const deletedAt = tracked.before.findIndex(({ id }) => id === deletedId);
  tracked.reviewer.acceptAll();
  direct.reviewer.acceptAll();
  const accepted = labelsOf(tracked.reviewer);
  const expected: string[] = [];
  const byGap = kept.toSorted((a, b) => a.gap - b.gap || a.at - b.at);
  for (const [index, block] of [...tracked.before, undefined].entries()) {
    for (const { at, gap } of byGap) if (gap === index) expected.push(`I${at}`);
    const label = block?.label;
    if (label !== undefined && label !== expected.at(-1))
      if (accepted.includes(label)) expected.push(label);
  }
  return {
    kept,
    tracked,
    direct,
    accepted,
    directAccepted: labelsOf(direct.reviewer),
    expected,
    // Deleting the paragraph a pending join runs words into takes those
    // words too only when tracked (`deletionTakesJoinedWords`), which is not
    // about where blocks land.
    deletesJoinedWords: tracked.markDeleted.has(
      tracked.before[deletedAt - 1]?.id ?? "",
    ),
  };
};

const caseArbitrary: fc.Arbitrary<Case> = fc
  .record({
    paragraphs: fc.integer({ min: 3, max: 4 }),
    table: fc.boolean(),
    setup: fc.constantFrom<Setup[]>(
      "none",
      "emptiedJoin",
      "trackedJoin",
      "finalDeleted",
    ),
    target: fc.integer({ min: 0, max: 2 }),
    insertions: fc.array(
      fc.record({
        kind: fc.constantFrom<Insertion["kind"][]>(
          "paragraph",
          "paragraph",
          "table",
        ),
        side: fc.constantFrom<Insertion["side"][]>("before", "after"),
        // A few anchors, so insertions share them.
        anchor: fc.integer({ min: 0, max: 4 }),
      }),
      { minLength: 2, maxLength: 3 },
    ),
    deletes: fc.option(fc.integer({ min: 0, max: 3 }), { nil: undefined }),
  })
  .map(({ paragraphs, table, setup, target, insertions, deletes }) => ({
    paragraphs,
    tableAfter: table ? 1 : undefined,
    setup,
    target,
    deletes,
    // A batch that deletes a paragraph puts paragraphs beside paragraphs:
    // tables a direct batch puts right beside the paragraph it deletes can
    // still leave it no paragraph to go by when it runs (the examples hold
    // the ones that do not).
    insertions:
      deletes === undefined
        ? insertions
        : insertions.map(({ side, anchor }) => ({
            kind: "paragraph" as const,
            side,
            anchor: anchor % paragraphs,
          })),
  }));

const paragraph = (side: Insertion["side"], anchor: number): Insertion => ({
  kind: "paragraph",
  side,
  anchor,
});
const table = (side: Insertion["side"], anchor: number): Insertion => ({
  kind: "table",
  side,
  anchor,
});
const plain = {
  paragraphs: 3,
  tableAfter: undefined,
  setup: "none",
  target: 0,
  deletes: undefined,
} as const;

const examples: [Case][] = [
  // An emptied paragraph whose join leaves the next one, and a table after
  // that one: a table and paragraphs after either keep the batch's order.
  [
    {
      ...plain,
      paragraphs: 4,
      setup: "emptiedJoin",
      target: 2,
      insertions: [
        table("after", 3),
        paragraph("after", 2),
        paragraph("after", 3),
      ],
    },
  ],
  // Tracked join: a paragraph before the joining paragraph after one put
  // after it stays before it.
  [
    {
      ...plain,
      setup: "trackedJoin",
      target: 1,
      insertions: [paragraph("after", 1), paragraph("before", 1)],
    },
  ],
  // Tracked join: a table after the joining paragraph follows the paragraph
  // the batch already put after it.
  [
    {
      ...plain,
      setup: "trackedJoin",
      insertions: [
        paragraph("after", 0),
        paragraph("before", 0),
        table("after", 0),
      ],
    },
  ],
  // Tracked join: a paragraph after the joining paragraph, then a table
  // before the one it joins, are both admitted.
  [
    {
      ...plain,
      setup: "trackedJoin",
      insertions: [paragraph("after", 0), table("before", 1)],
    },
  ],
  // The last paragraph deleted: a paragraph before it ends the join for one
  // put after the paragraph before it later in the batch.
  [
    {
      ...plain,
      setup: "finalDeleted",
      insertions: [
        paragraph("after", 1),
        paragraph("before", 2),
        paragraph("after", 1),
      ],
    },
  ],
  // Direct: a paragraph after an emptied paragraph lands after the one its
  // join leaves, and the last paragraph's deletion still follows it.
  [
    {
      ...plain,
      setup: "emptiedJoin",
      insertions: [paragraph("before", 0), paragraph("after", 0)],
      deletes: 2,
    },
  ],
  // Direct: a table put before the paragraph a join leaves goes before the
  // join, whose first paragraph the batch deletes after it lands.
  [
    {
      ...plain,
      setup: "trackedJoin",
      insertions: [paragraph("before", 0), table("before", 1)],
      deletes: 0,
    },
  ],
  // Direct: a paragraph deleted between a table and one the batch puts
  // before the next paragraph keeps the two tables apart.
  [
    {
      ...plain,
      paragraphs: 4,
      tableAfter: 1,
      insertions: [paragraph("before", 0), table("before", 3)],
      deletes: 2,
    },
  ],
  // Direct: blocks after a paragraph and before the next, with the one
  // between them deleted, stay on their sides.
  [
    {
      ...plain,
      insertions: [
        paragraph("before", 2),
        paragraph("after", 0),
        paragraph("after", 0),
      ],
      deletes: 1,
    },
  ],
  // An emptied paragraph before a table: a paragraph and a table after it
  // are both admitted, and a block put before the table after the paragraph
  // the batch put there follows that one.
  [
    {
      ...plain,
      setup: "emptiedJoin",
      target: 1,
      tableAfter: 1,
      insertions: [paragraph("after", 1), table("after", 1)],
    },
  ],
  [
    {
      ...plain,
      setup: "emptiedJoin",
      target: 1,
      tableAfter: 1,
      insertions: [paragraph("after", 1), paragraph("before", 3)],
    },
  ],
  // Direct: tables put after the paragraph before a deleted one, and right
  // after it or before the paragraph after it, leave it a paragraph to go
  // by: the batch's paragraph before it goes after it.
  [
    {
      ...plain,
      paragraphs: 4,
      setup: "emptiedJoin",
      insertions: [
        paragraph("before", 2),
        table("after", 0),
        table("before", 3),
      ],
      deletes: 2,
    },
  ],
  [
    {
      ...plain,
      insertions: [
        paragraph("before", 1),
        table("after", 1),
        table("after", 0),
      ],
      deletes: 1,
    },
  ],
  // Tracked: a table after a table and a paragraph before the deleted last
  // paragraph after it take revision ids of their own as they replay.
  [
    {
      ...plain,
      setup: "finalDeleted",
      tableAfter: 1,
      insertions: [table("after", 3), paragraph("before", 2)],
      deletes: 2,
    },
  ],
  // Direct: a paragraph before a deleted one a table follows, and one after
  // the paragraph before it, keep the batch's order past it.
  [
    {
      ...plain,
      tableAfter: 1,
      insertions: [paragraph("before", 1), paragraph("after", 0)],
      deletes: 1,
    },
  ],
];

const PROPERTY =
  "a batch puts blocks in the gaps they land in one at a time, in batch order, tracked and direct alike";

test(
  PROPERTY,
  async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async (setup) => {
        const reviewed = await replay(setup);
        fc.pre(
          reviewed.kept.length >= 2 &&
            !reviewed.tracked.skipped.includes("delete") &&
            !reviewed.deletesJoinedWords,
        );
        expect(reviewed.tracked.skipped).toEqual([]);
        expect(reviewed.direct.skipped).toEqual([]);
        expect(reviewed.directAccepted).toEqual(reviewed.accepted);
        expect(reviewed.accepted).toEqual(reviewed.expected);
      }),
      { id: PROPERTY, numRuns: 60, examples },
    );
  },
  propertyTestTimeout(120_000),
);

// The reported case: the words of the paragraph before the last one were
// deleted and accepted, so it stands emptied with its break pending deletion,
// read as one paragraph with the last one. Put after it, a paragraph lands
// after the last one, and one put after the last one afterwards follows it.
test.each(["tracked-changes", "direct"] as const)(
  "blocks after an emptied paragraph and after the one its join leaves keep the batch's order (%s)",
  async (mode) => {
    const setup: Case = {
      ...plain,
      paragraphs: 4,
      setup: "emptiedJoin",
      target: 2,
      insertions: [paragraph("after", 2), paragraph("after", 3)],
    };
    const reviewer = await prepared(setup);
    const anchors = anchorsOf(reviewer);
    const result = reviewer.applyOperations(
      setup.insertions.map((insertion, at) =>
        operationOf(insertion, at, anchors),
      ),
      { mode },
    );
    expect(result.skipped).toEqual([]);
    reviewer.acceptAll();
    expect(labelsOf(reviewer)).toEqual(["P0", "P1", "P3", "I0", "I1"]);
  },
);
