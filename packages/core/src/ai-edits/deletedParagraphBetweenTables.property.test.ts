/**
 * A paragraph a batch deletes beside a table reads the same applied directly
 * and as tracked changes accepted, once saved and reopened; rejecting every
 * tracked change gives the original back. Between two tables it stays,
 * emptied, and keeps them apart; before or after one it goes whole.
 *
 * The block before the deleted paragraph is a paragraph, which may get a table
 * put after it, or a table; the block after it is nothing, a paragraph or a
 * table. An earlier batch may put tables and paragraphs beside it first, so
 * the later one meets them pending in the tracked replay. The batch deletes
 * the paragraph, in any place among up to three tables or paragraphs it puts
 * beside it or before the block that follows it (a table's cell names the
 * table). The blocks stand in the body, a text box or a table cell.
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
  /** The deleted paragraph, or the block after it as the batch finds it. */
  target: "anchor" | "follower";
};
type Case = {
  container: "body" | "textBox" | "cell";
  before: "paragraph" | "table";
  after: "end" | "paragraph" | "table";
  /** A table put after the paragraph before, when that is a paragraph. */
  tableAfterPrevious: boolean;
  /** An earlier batch's insertions beside the paragraph. */
  earlier: Insertion[];
  insertions: Insertion[];
  /** Where the deletion stands among the batch's operations. */
  deletionAt: number;
};

const OPENING = "Opening clause.";
const PREVIOUS = "Previous clause.";
const ANCHOR = "Anchor clause.";
const NEXT = "Next clause.";
const CLOSING = "Closing clause.";

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
  const block = reviewer.getContent().find((entry) => entry.text === text);
  if (block === undefined) throw new Error(`Missing block ${text}`);
  return block.id;
};

/** The block after the anchor as the batch finds it, the anchor at the end. */
const followerOf = (reviewer: FolioDocxReviewer, anchor: string): string => {
  const content = reviewer.getContent();
  const index = content.findIndex(({ id }) => id === anchor);
  return content[index + 1]?.id ?? anchor;
};

const insertionsOf = (
  insertions: readonly Insertion[],
  reviewer: FolioDocxReviewer,
  prefix: string,
): FolioDocumentOperation[] => {
  const anchor = blockIdOf(reviewer, ANCHOR);
  const follower = followerOf(reviewer, anchor);
  return insertions.map(({ kind, side, target }, index) => {
    const id = `${prefix}-${String(index)}`;
    const blockId = target === "anchor" ? anchor : follower;
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
): FolioDocumentOperation[] => {
  const operations = insertionsOf(setup.insertions, reviewer, "insert");
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
    blockId: blockIdOf(reviewer, ANCHOR),
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
  for (const operations of [
    () => insertionsOf(setup.earlier, reviewer, "earlier"),
    () => operationsOf(setup, reviewer),
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
  target: fc.constantFrom<Insertion["target"][]>("anchor", "follower"),
});

const caseArbitrary: fc.Arbitrary<Case> = fc
  .record({
    container: fc.constantFrom<Case["container"][]>("body", "textBox", "cell"),
    before: fc.constantFrom<Case["before"][]>("paragraph", "table"),
    after: fc.constantFrom<Case["after"][]>("end", "paragraph", "table"),
    tableAfterPrevious: fc.boolean(),
    earlier: fc.array(insertionArbitrary, { maxLength: 2 }),
    insertions: fc.array(insertionArbitrary, { maxLength: 3 }),
    deletionAt: fc.nat({ max: 4 }),
  })
  .map((setup) =>
    Object.assign(setup, {
      // Only a paragraph before it can have a table put after it.
      tableAfterPrevious:
        setup.before === "paragraph" && setup.tableAfterPrevious,
    }),
  );

const anchorInsertion = (
  kind: Insertion["kind"],
  side: Insertion["side"],
): Insertion => ({ kind, side, target: "anchor" });

const singleBatch = {
  container: "body",
  tableAfterPrevious: false,
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
      earlier: [],
      insertions: [anchorInsertion("paragraph", "before")],
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
