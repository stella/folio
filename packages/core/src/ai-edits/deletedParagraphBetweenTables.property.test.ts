/**
 * A paragraph a batch deletes that would then stand between two tables stays,
 * emptied, and keeps them apart: applied directly, and as tracked changes
 * accepted, which read the same once saved and reopened.
 *
 * The block before the deleted paragraph is a paragraph, which may get a table
 * put after it, or a table; the block after it is nothing, a paragraph or a
 * table. The batch deletes the paragraph, in any place among up to three
 * blocks, tables or paragraphs, it puts before or after it.
 *
 * A case applies a batch to a fresh package in each mode, then saves and
 * reopens it, about 200 ms; the budget is 60 s for the stated runs, scaled by
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
import { FolioDocxReviewer } from "./headless";

type Mode = "direct" | "tracked-changes";
type Insertion = { kind: "table" | "paragraph"; side: "before" | "after" };
type Case = {
  before: "paragraph" | "table";
  after: "end" | "paragraph" | "table";
  /** A table put after the paragraph before, when that is a paragraph. */
  tableAfterPrevious: boolean;
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

const open = async ({ before, after }: Case): Promise<FolioDocxReviewer> => {
  const texts = [
    OPENING,
    before === "table" ? tableMarkdown("Before") : PREVIOUS,
    ANCHOR,
    ...(after === "paragraph" ? [NEXT] : []),
    ...(after === "table" ? [tableMarkdown("After"), CLOSING] : []),
  ];
  const bytes = (
    await ensureParaIds(await createDocx(fromMarkdown(texts.join("\n\n"))))
  ).docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
};

const blockIdOf = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((entry) => entry.text === text);
  if (block === undefined) throw new Error(`Missing block ${text}`);
  return block.id;
};

const operationsOf = (
  setup: Case,
  reviewer: FolioDocxReviewer,
): FolioDocumentOperation[] => {
  const anchor = blockIdOf(reviewer, ANCHOR);
  const operations: FolioDocumentOperation[] = setup.insertions.map(
    ({ kind, side }, index) =>
      kind === "table"
        ? {
            id: `insert-${String(index)}`,
            type: "insertTable",
            blockId: anchor,
            position: side,
            rows: [[`Inserted ${String(index)}`]],
          }
        : {
            id: `insert-${String(index)}`,
            type: side === "before" ? "insertBeforeBlock" : "insertAfterBlock",
            blockId: anchor,
            text: `Inserted ${String(index)}.`,
          },
  );
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
    blockId: anchor,
  });
  return operations;
};

/** The blocks a reader lists once the batch is applied, accepted and saved. */
const replay = async (setup: Case, mode: Mode) => {
  const reviewer = await open(setup);
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      atomic: false,
      operations: operationsOf(setup, reviewer),
    },
    { undefinedReferences: "refuse" },
  );
  reviewer.acceptAll();
  const saved = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
    author: "Reviewer",
  });
  return {
    applied: result.applied.map(({ id }) => id).toSorted(),
    blocks: saved.getContent().map(({ kind, text, table }) => ({
      kind,
      text,
      table: table ?? null,
    })),
  };
};

const caseArbitrary: fc.Arbitrary<Case> = fc
  .record({
    before: fc.constantFrom<Case["before"][]>("paragraph", "table"),
    after: fc.constantFrom<Case["after"][]>("end", "paragraph", "table"),
    tableAfterPrevious: fc.boolean(),
    insertions: fc.array(
      fc.record({
        kind: fc.constantFrom<Insertion["kind"][]>("table", "paragraph"),
        side: fc.constantFrom<Insertion["side"][]>("before", "after"),
      }),
      { maxLength: 3 },
    ),
    deletionAt: fc.nat({ max: 4 }),
  })
  .map(({ before, after, tableAfterPrevious, insertions, deletionAt }) => ({
    before,
    after,
    // Only a paragraph before it can have a table put after it.
    tableAfterPrevious: before === "paragraph" && tableAfterPrevious,
    insertions,
    deletionAt,
  }));

const examples: [Case][] = [
  // #1350 flow 905f2b7afc8ccdbe: the story's last paragraph after a table,
  // deleted, with a table put after it.
  [
    {
      before: "table",
      after: "end",
      tableAfterPrevious: false,
      insertions: [{ kind: "table", side: "after" }],
      deletionAt: 0,
    },
  ],
  // #1350 flow 705c1bae88cd5e7f: a table put after the paragraph before it
  // and one put after it.
  [
    {
      before: "paragraph",
      after: "paragraph",
      tableAfterPrevious: true,
      insertions: [{ kind: "table", side: "after" }],
      deletionAt: 1,
    },
  ],
];

test(
  "a deleted paragraph between two tables reads the same applied directly and accepted",
  async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async (setup) => {
        const direct = await replay(setup, "direct");
        const tracked = await replay(setup, "tracked-changes");
        // Where one mode refuses what the other applies, the two documents
        // part, and comparing them says nothing about the rule.
        fc.pre(
          JSON.stringify(direct.applied) === JSON.stringify(tracked.applied),
        );
        expect(tracked.blocks).toEqual(direct.blocks);
      }),
      {
        id: "a deleted paragraph between two tables reads the same applied directly and accepted",
        numRuns: 60,
        examples,
      },
    );
  },
  propertyTestTimeout(60_000),
);
