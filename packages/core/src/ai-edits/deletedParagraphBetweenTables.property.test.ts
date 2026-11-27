/**
 * A paragraph a batch deletes beside a table reads the same applied directly
 * and as tracked changes accepted, once saved and reopened; rejecting every
 * tracked change gives the original back. Between two tables it stays,
 * emptied, and keeps them apart; before or after one it goes whole.
 *
 * The block before the deleted paragraph is a paragraph, which may get a table
 * put after it, or a table; the block after it is nothing, a paragraph or a
 * table. A tracked join may be pending, of the opening paragraph into the
 * next, of the paragraph before into the deleted one, or of the deleted one
 * into the next, and an earlier batch may put tables and paragraphs beside it
 * first, so the later one meets them pending in the tracked replay. The
 * batch deletes the paragraph, in any place among up to three tables or
 * paragraphs it puts beside it, beside the blocks before and after it (a
 * table's cell names the table) or beside the opening paragraph. The blocks
 * stand in the body, a text box or a table cell.
 *
 * A case applies the batches to a fresh package in each mode, then saves and
 * reopens it, about 300 ms; the budget is 90 s for the stated runs, scaled by
 * the nightly factor.
 */

import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { BlockContent, Paragraph } from "../types/document";
import { FolioDocxReviewer } from "./headless";

type Mode = "direct" | "tracked-changes";
type Insertion = {
  kind: "table" | "paragraph";
  side: "before" | "after";
  /**
   * The deleted paragraph, the block after or before it as the batch finds
   * it (a table's cell, next to a table), or the opening paragraph.
   */
  target: "anchor" | "follower" | "precursor" | "opening";
};
type Case = {
  container: "body" | "textBox" | "cell";
  before: "paragraph" | "table";
  after: "end" | "paragraph" | "table";
  /** A table put after the paragraph before, when that is a paragraph. */
  tableAfterPrevious: boolean;
  /**
   * The paragraph a tracked join runs into the one after it, pending before
   * the batches in both replays: the opening one, the one before the deleted
   * paragraph, which runs into it, or the deleted one, which runs out of it,
   * when a paragraph follows each.
   */
  pendingJoin: "none" | "opening" | "previous" | "anchor";
  /** An earlier batch's insertions beside the paragraph. */
  earlier: readonly Insertion[];
  insertions: readonly Insertion[];
  /** Where the deletion stands among the batch's operations. */
  deletionAt: number;
};

const OPENING = "Opening clause.";
const PREVIOUS = "Previous clause.";
const ANCHOR = "Anchor clause.";
const NEXT = "Next clause.";
const CLOSING = "Closing clause.";
const JOINED = { opening: OPENING, previous: PREVIOUS, anchor: ANCHOR };

const tableMarkdown = (label: string) =>
  `| ${label} one | ${label} two |\n| --- | --- |\n| ${label} three | ${label} four |`;

const textParagraph = (text: string): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

/** `blocks` where the case puts them: the body, a text box or a cell. */
const contained = (
  container: Case["container"],
  blocks: BlockContent[],
): BlockContent[] => {
  switch (container) {
    case "body":
      return blocks;
    case "textBox":
      return [
        {
          type: "paragraph",
          content: [
            {
              type: "run",
              content: [
                {
                  type: "shape",
                  shape: {
                    type: "shape",
                    shapeType: "textBox",
                    id: "case-text-box",
                    size: { width: 2_000_000, height: 800_000 },
                    textBody: { content: blocks },
                  },
                },
              ],
            },
          ],
        },
        textParagraph("Body clause."),
      ];
    case "cell":
      return [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [{ type: "tableCell", content: blocks }],
            },
          ],
        },
        textParagraph("Body clause."),
      ];
  }
};

const open = async ({
  container,
  before,
  after,
}: Case): Promise<FolioDocxReviewer> => {
  const texts = [
    OPENING,
    before === "table" ? tableMarkdown("Before") : PREVIOUS,
    ANCHOR,
    ...(after === "paragraph" ? [NEXT] : []),
    ...(after === "table" ? [tableMarkdown("After"), CLOSING] : []),
  ];
  const document = fromMarkdown(texts.join("\n\n"));
  document.package.document.content = contained(
    container,
    document.package.document.content,
  );
  const bytes = (await ensureParaIds(await createDocx(document))).docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
};

const blockIdOf = (reviewer: FolioDocxReviewer, text: string): string => {
  // A pending join reads the opening paragraph with the one it joins.
  const block = reviewer
    .getContent()
    .find((entry) => entry.text.startsWith(text));
  if (block === undefined) throw new Error(`Missing block ${text}`);
  return block.id;
};

/**
 * The deleted paragraph, which keeps its id through a join into it, and the
 * paragraph a pending join runs into it, if any, which the direct join
 * removes: it reads as part of the deleted paragraph.
 */
type Anchor = { id: string; joined: string | undefined };

/** The block `step` places from the anchor as the batch finds it, or the anchor. */
const besideOf = (
  reviewer: FolioDocxReviewer,
  { id: anchor, joined }: Anchor,
  step: 1 | -1,
): string => {
  const content = reviewer.getContent().filter(({ id }) => id !== joined);
  const index = content.findIndex(({ id }) => id === anchor);
  return content[index + step]?.id ?? anchor;
};

const insertionsOf = (
  insertions: readonly Insertion[],
  reviewer: FolioDocxReviewer,
  anchor: Anchor,
  prefix: string,
): FolioDocumentOperation[] => {
  const targets = {
    anchor: anchor.id,
    follower: besideOf(reviewer, anchor, 1),
    precursor: besideOf(reviewer, anchor, -1),
    opening: blockIdOf(reviewer, OPENING),
  };
  return insertions.map(({ kind, side, target }, index) => {
    const id = `${prefix}-${String(index)}`;
    const blockId = targets[target];
    const text = `Inserted ${prefix} ${String(index)}`;
    return kind === "table"
      ? { id, type: "insertTable", blockId, position: side, rows: [[text]] }
      : {
          id,
          type: side === "before" ? "insertBeforeBlock" : "insertAfterBlock",
          blockId,
          text: `${text}.`,
        };
  });
};

const operationsOf = (
  setup: Case,
  reviewer: FolioDocxReviewer,
  anchor: Anchor,
): FolioDocumentOperation[] => {
  const operations = insertionsOf(setup.insertions, reviewer, anchor, "insert");
  if (setup.tableAfterPrevious)
    operations.unshift({
      id: "table-after-previous",
      type: "insertTable",
      blockId: blockIdOf(reviewer, PREVIOUS),
      rows: [["After previous"]],
    });
  operations.splice(Math.min(setup.deletionAt, operations.length), 0, {
    id: "delete",
    type: "deleteBlock",
    blockId: anchor.id,
  });
  return operations;
};

const blocksOf = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ kind, text, table }) => ({
    kind,
    text,
    table: table ?? null,
  }));

const reopen = async (reviewer: FolioDocxReviewer) =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
    author: "Reviewer",
  });

/**
 * The blocks a reader lists once the batches are applied, accepted and
 * saved; tracked, also once they are rejected.
 */
const replay = async (setup: Case, mode: Mode) => {
  const reviewer = await open(setup);
  const original = blocksOf(reviewer);
  const applied: string[] = [];
  const anchor: Anchor = {
    id: blockIdOf(reviewer, ANCHOR),
    joined:
      setup.pendingJoin === "previous"
        ? blockIdOf(reviewer, PREVIOUS)
        : undefined,
  };
  if (setup.pendingJoin !== "none") {
    const joined = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        // A join into the deleted paragraph goes in the batches' mode: the
        // tracked deletion takes the words it runs in, as deleting the
        // joined paragraph directly does.
        mode: setup.pendingJoin === "previous" ? mode : "tracked-changes",
        operations: [
          {
            id: "join",
            type: "mergeBlockWithNext",
            blockId: blockIdOf(reviewer, JOINED[setup.pendingJoin]),
            separator: " ",
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    applied.push(...joined.applied.map(({ id }) => id));
  }
  for (const operations of [
    () => insertionsOf(setup.earlier, reviewer, anchor, "earlier"),
    () => operationsOf(setup, reviewer, anchor),
  ]) {
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        atomic: false,
        operations: operations(),
      },
      { undefinedReferences: "refuse" },
    );
    applied.push(...result.applied.map(({ id }) => id).toSorted());
  }
  const rejected = await reopen(reviewer);
  rejected.rejectAll();
  reviewer.acceptAll();
  return {
    applied,
    original,
    blocks: blocksOf(await reopen(reviewer)),
    rejected: blocksOf(await reopen(rejected)),
  };
};

const insertionArbitrary: fc.Arbitrary<Insertion> = fc.record({
  kind: fc.constantFrom<Insertion["kind"][]>("table", "paragraph"),
  side: fc.constantFrom<Insertion["side"][]>("before", "after"),
  target: fc.constantFrom<Insertion["target"][]>(
    "anchor",
    "follower",
    "precursor",
    "opening",
  ),
});

const caseArbitrary: fc.Arbitrary<Case> = fc
  .record({
    container: fc.constantFrom<Case["container"][]>("body", "textBox", "cell"),
    before: fc.constantFrom<Case["before"][]>("paragraph", "table"),
    after: fc.constantFrom<Case["after"][]>("end", "paragraph", "table"),
    tableAfterPrevious: fc.boolean(),
    pendingJoin: fc.constantFrom<Case["pendingJoin"][]>(
      "none",
      "opening",
      "previous",
      "anchor",
    ),
    earlier: fc.array(insertionArbitrary, { maxLength: 2 }),
    insertions: fc.array(insertionArbitrary, { maxLength: 3 }),
    deletionAt: fc.nat({ max: 4 }),
  })
  .map((setup) =>
    Object.assign(setup, {
      // Only a paragraph before it can have a table put after it, one a
      // direct join leaves standing.
      tableAfterPrevious:
        setup.before === "paragraph" &&
        setup.pendingJoin !== "previous" &&
        setup.tableAfterPrevious,
      // Only a paragraph after the one a pending join runs out of takes it.
      pendingJoin: (
        setup.pendingJoin === "anchor"
          ? setup.after === "paragraph"
          : setup.before === "paragraph"
      )
        ? setup.pendingJoin
        : "none",
    }),
  );

const anchorInsertion = (
  kind: Insertion["kind"],
  side: Insertion["side"],
): Insertion => ({ kind, side, target: "anchor" });

const singleBatch = {
  container: "body",
  tableAfterPrevious: false,
  pendingJoin: "none",
  earlier: [],
} as const;

const regressions: [string, Case][] = [
  // #1350 flow 889abb90c8bd2d4f.
  [
    "a text box's paragraph with a pending paragraph after it, deleted as a table goes after it",
    {
      container: "textBox",
      before: "paragraph",
      after: "end",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [anchorInsertion("paragraph", "after")],
      insertions: [anchorInsertion("table", "after")],
      deletionAt: 0,
    },
  ],
  // #1350 flow d2d995429b7408e2.
  [
    "the body's last paragraph with a pending table after it, deleted as a paragraph goes before that table",
    {
      container: "body",
      before: "paragraph",
      after: "end",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [anchorInsertion("table", "after")],
      insertions: [{ kind: "paragraph", side: "before", target: "follower" }],
      deletionAt: 0,
    },
  ],
  [
    "the body's last paragraph with a pending table after it, deleted as a paragraph and a table go before that table",
    {
      container: "body",
      before: "paragraph",
      after: "end",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [anchorInsertion("table", "after")],
      insertions: [
        { kind: "paragraph", side: "before", target: "follower" },
        { kind: "table", side: "before", target: "follower" },
      ],
      deletionAt: 0,
    },
  ],
  [
    "a cell's paragraph before a nested table, deleted as a paragraph goes before it",
    {
      container: "cell",
      before: "paragraph",
      after: "table",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [],
      insertions: [anchorInsertion("paragraph", "before")],
      deletionAt: 0,
    },
  ],
  // Seed 1758728747: what is put beside a cell's paragraph lands beside its
  // table, which the deletion inside it leaves where it was put.
  [
    "a cell's paragraph before a nested table, deleted as paragraphs go before and after it and a table before it",
    {
      container: "cell",
      before: "paragraph",
      after: "table",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [],
      insertions: [
        anchorInsertion("paragraph", "before"),
        anchorInsertion("paragraph", "after"),
        anchorInsertion("table", "before"),
      ],
      deletionAt: 0,
    },
  ],
  // A table put after the paragraph a pending join takes lands after the
  // join, right before the deleted paragraph, which then stands between two
  // tables.
  [
    "the body's last paragraph after a pending join, deleted as a paragraph goes before it and tables after the join and after it",
    {
      container: "body",
      before: "paragraph",
      after: "end",
      tableAfterPrevious: false,
      pendingJoin: "opening",
      earlier: [],
      insertions: [
        anchorInsertion("paragraph", "before"),
        { kind: "table", side: "after", target: "opening" },
        anchorInsertion("table", "after"),
      ],
      deletionAt: 3,
    },
  ],
  // A table put before the next table's cell lands right after the deleted
  // paragraph, which then stands between two tables.
  [
    "a paragraph before a table, deleted as tables go before it and before that table's cell and a paragraph after it",
    {
      container: "body",
      before: "paragraph",
      after: "table",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [],
      insertions: [
        anchorInsertion("table", "before"),
        { kind: "table", side: "before", target: "follower" },
        anchorInsertion("paragraph", "after"),
      ],
      deletionAt: 3,
    },
  ],
  // Put after the cell's paragraph before it, a table and a paragraph land
  // after the cell's table, not beside the deleted paragraph.
  [
    "a cell's paragraph before a nested table, deleted as a table and a paragraph go after the cell paragraph before it",
    {
      container: "cell",
      before: "paragraph",
      after: "table",
      tableAfterPrevious: true,
      pendingJoin: "none",
      earlier: [],
      insertions: [{ kind: "paragraph", side: "after", target: "precursor" }],
      deletionAt: 0,
    },
  ],
  // Directly, the tables put after the paragraph before it stay ahead of the
  // paragraph put after it, which stands in its place.
  [
    "a paragraph deleted as two tables go after the one before it and a paragraph after it",
    {
      container: "body",
      before: "paragraph",
      after: "paragraph",
      tableAfterPrevious: true,
      pendingJoin: "none",
      earlier: [],
      insertions: [
        anchorInsertion("paragraph", "after"),
        { kind: "table", side: "after", target: "precursor" },
      ],
      deletionAt: 0,
    },
  ],
  // A paragraph put after the paragraph before it, after a table put there,
  // stands right before it: it goes whole, not kept between tables.
  [
    "a paragraph before a table, deleted as a table and then a paragraph go after the one before it",
    {
      container: "body",
      before: "paragraph",
      after: "table",
      tableAfterPrevious: true,
      pendingJoin: "none",
      earlier: [],
      insertions: [{ kind: "paragraph", side: "after", target: "precursor" }],
      deletionAt: 0,
    },
  ],
  // Put after the paragraph before it later in the batch, the table lands
  // after the paragraph put before it, so it stands between two tables.
  [
    "a paragraph before a table, deleted as a paragraph goes before it and then a table after the one before it",
    {
      container: "body",
      before: "paragraph",
      after: "table",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [],
      insertions: [
        anchorInsertion("paragraph", "before"),
        { kind: "table", side: "after", target: "precursor" },
      ],
      deletionAt: 0,
    },
  ],
  // Directly, both go after it before it goes, in batch order.
  [
    "a paragraph before a table, deleted as paragraphs go before it and then after the one before it",
    {
      container: "body",
      before: "paragraph",
      after: "table",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [],
      insertions: [
        anchorInsertion("paragraph", "before"),
        { kind: "paragraph", side: "after", target: "precursor" },
      ],
      deletionAt: 0,
    },
  ],
  // Directly, the tables put after the paragraph before it stay ahead of the
  // paragraph put before the one after it.
  [
    "a paragraph deleted as two tables go after the one before it and a paragraph before the next",
    {
      container: "body",
      before: "paragraph",
      after: "paragraph",
      tableAfterPrevious: true,
      pendingJoin: "none",
      earlier: [],
      insertions: [
        { kind: "table", side: "after", target: "precursor" },
        { kind: "paragraph", side: "before", target: "follower" },
      ],
      deletionAt: 0,
    },
  ],
  // Put after the table before it, named by that table's cell, a table lands
  // in its gap after the paragraph put before it, in batch order.
  [
    "a paragraph after a table, deleted as a paragraph goes before it and a table after that table's cell",
    {
      container: "body",
      before: "table",
      after: "paragraph",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [],
      insertions: [
        anchorInsertion("paragraph", "before"),
        { kind: "table", side: "after", target: "precursor" },
      ],
      deletionAt: 0,
    },
  ],
  [
    "the body's last paragraph after a table, deleted as a paragraph goes before it and tables after it and after that table's cell",
    {
      container: "body",
      before: "table",
      after: "end",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [],
      insertions: [
        anchorInsertion("paragraph", "before"),
        anchorInsertion("table", "after"),
        { kind: "table", side: "after", target: "precursor" },
      ],
      deletionAt: 0,
    },
  ],
  // Directly, tables put after a pending table's cell replay beside that
  // table, not by the paragraph after it, which the batch removes first.
  [
    "a paragraph after a pending table, deleted as two tables go after that table's cell",
    {
      container: "body",
      before: "paragraph",
      after: "paragraph",
      tableAfterPrevious: false,
      pendingJoin: "none",
      earlier: [{ kind: "table", side: "after", target: "precursor" }],
      insertions: [
        { kind: "table", side: "after", target: "precursor" },
        { kind: "table", side: "after", target: "precursor" },
      ],
      deletionAt: 0,
    },
  ],
  [
    "a paragraph after a pending join, deleted between tables put after the join and before the next paragraph",
    {
      container: "body",
      before: "paragraph",
      after: "paragraph",
      tableAfterPrevious: true,
      pendingJoin: "opening",
      earlier: [],
      insertions: [
        { kind: "table", side: "before", target: "follower" },
        { kind: "table", side: "after", target: "opening" },
      ],
      deletionAt: 0,
    },
  ],
  // #3690 flow a0ac291309b51011: the table lands right before the join,
  // so the paragraph follows a table and keeps its mark, its words and the
  // joined ones going.
  [
    "the body's last paragraph a pending join runs into, deleted as a table goes after the paragraph before the join",
    {
      ...singleBatch,
      before: "paragraph",
      after: "end",
      pendingJoin: "previous",
      insertions: [{ kind: "table", side: "after", target: "opening" }],
      deletionAt: 1,
    },
  ],
  // Put before a paragraph the batch deletes, a table lands after the one
  // before the join that runs into it, not after the join.
  [
    "the body's last paragraph a pending join runs into, deleted as a table goes before it",
    {
      ...singleBatch,
      before: "paragraph",
      after: "end",
      pendingJoin: "previous",
      insertions: [anchorInsertion("table", "before")],
      deletionAt: 0,
    },
  ],
  [
    "a paragraph a pending join runs into, deleted between a table after it and one put before it",
    {
      ...singleBatch,
      before: "paragraph",
      after: "table",
      pendingJoin: "previous",
      insertions: [anchorInsertion("table", "before")],
      deletionAt: 0,
    },
  ],
  [
    "a paragraph a pending join runs into, deleted between a table after it and one put after the paragraph before the join",
    {
      ...singleBatch,
      before: "paragraph",
      after: "table",
      pendingJoin: "previous",
      insertions: [{ kind: "table", side: "after", target: "opening" }],
      deletionAt: 0,
    },
  ],
  [
    "a paragraph a pending join runs into, deleted between tables put before and after it",
    {
      ...singleBatch,
      before: "paragraph",
      after: "paragraph",
      pendingJoin: "previous",
      insertions: [
        anchorInsertion("table", "before"),
        anchorInsertion("table", "after"),
      ],
      deletionAt: 0,
    },
  ],
  // A paragraph put right before the join stands before it when the
  // deletion runs, which retires that one's break and takes the joined words.
  [
    "the body's last paragraph a pending join runs into, deleted as a paragraph goes after the one before the join",
    {
      ...singleBatch,
      before: "paragraph",
      after: "end",
      pendingJoin: "previous",
      insertions: [{ kind: "paragraph", side: "after", target: "precursor" }],
      deletionAt: 0,
    },
  ],
  [
    "the body's last paragraph a pending join runs into, deleted as a table and then a paragraph go before it",
    {
      ...singleBatch,
      before: "paragraph",
      after: "end",
      pendingJoin: "previous",
      insertions: [
        anchorInsertion("table", "before"),
        anchorInsertion("paragraph", "before"),
      ],
      deletionAt: 0,
    },
  ],
  [
    "a paragraph a pending join runs into before a table, deleted as a paragraph goes after a pending table before the join",
    {
      ...singleBatch,
      before: "paragraph",
      after: "table",
      pendingJoin: "previous",
      earlier: [anchorInsertion("table", "before")],
      insertions: [{ kind: "paragraph", side: "after", target: "precursor" }],
      deletionAt: 0,
    },
  ],
  // A paragraph and a table put after a table's cell land once the deletion
  // has run: the deleted paragraph still stands between that table and the
  // one after it, and directly the table keeps its side of it too.
  [
    "a paragraph between a pending table and a table, deleted as a paragraph and a table go after the first table's cell",
    {
      ...singleBatch,
      before: "paragraph",
      after: "table",
      earlier: [anchorInsertion("table", "before")],
      insertions: [
        { kind: "paragraph", side: "after", target: "precursor" },
        { kind: "table", side: "after", target: "precursor" },
      ],
      deletionAt: 0,
    },
  ],
];

const examples: [Case][] = [
  // #1350 flow 905f2b7afc8ccdbe: the story's last paragraph after a table,
  // deleted, with a table put after it.
  [
    {
      ...singleBatch,
      before: "table",
      after: "end",
      insertions: [anchorInsertion("table", "after")],
      deletionAt: 0,
    },
  ],
  // #1350 flow 705c1bae88cd5e7f: a table put after the paragraph before it
  // and one put after it.
  [
    {
      ...singleBatch,
      before: "paragraph",
      after: "paragraph",
      tableAfterPrevious: true,
      insertions: [anchorInsertion("table", "after")],
      deletionAt: 1,
    },
  ],
  ...regressions.map(([, setup]): [Case] => [setup]),
];

test.each(regressions)(
  "%s: accepted, it goes as it does directly",
  async (_name, setup) => {
    const direct = await replay(setup, "direct");
    const tracked = await replay(setup, "tracked-changes");
    expect(tracked.applied).toEqual(direct.applied);
    expect(direct.blocks.map(({ text }) => text)).not.toContain(ANCHOR);
    expect(tracked.blocks).toEqual(direct.blocks);
    expect(tracked.rejected).toEqual(tracked.original);
  },
);

test(
  "a deleted paragraph beside a table reads the same applied directly and accepted",
  async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async (setup) => {
        const direct = await replay(setup, "direct");
        const tracked = await replay(setup, "tracked-changes");
        expect(tracked.rejected).toEqual(tracked.original);
        // Where one mode refuses what the other applies, the two documents
        // part, and comparing them says nothing about the rule.
        fc.pre(
          JSON.stringify(direct.applied) === JSON.stringify(tracked.applied),
        );
        expect(tracked.blocks).toEqual(direct.blocks);
      }),
      {
        id: "a deleted paragraph beside a table reads the same applied directly and accepted",
        numRuns: 60,
        examples,
      },
    );
  },
  propertyTestTimeout(90_000),
);
