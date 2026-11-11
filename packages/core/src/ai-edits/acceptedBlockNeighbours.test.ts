/**
 * An edit reads a document with pending changes as accepting leaves it, past
 * the blocks accepting removes whole: applied as tracked changes and accepted,
 * it reads as the same edit applied directly.
 *
 * - A merge whose next paragraphs accepting removes, up to a table that
 *   stays, has no paragraph to join, as right before a table: it is refused in
 *   every mode, rather than running the paragraph into the table's first cell
 *   directly while the tracked one keeps it apart.
 * - Blocks put on either side of a table accepting removes share one gap and
 *   keep the order the batch names them in, as they do once the table is
 *   removed directly.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

type Mode = "direct" | "tracked-changes" | "suggested";

const TABLE = "| Item | Price |\n| --- | --- |\n| Widget | 10 |";
const ROWS = [
  ["Term", "Value"],
  ["Period", "12 months"],
];

const open = async (markdown: string): Promise<FolioDocxReviewer> => {
  const bytes = (await ensureParaIds(await createDocx(fromMarkdown(markdown))))
    .docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Test" });
};

const idOf = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.snapshot().blocks.find((b) => b.text === text);
  if (!block) throw new Error(`No block reads "${text}"`);
  return block.id;
};

const apply = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      atomic: false,
      operations,
    },
    { undefinedReferences: "refuse" },
  );

/** The blocks once every change is accepted, saved and reopened. */
const accepted = async (reviewer: FolioDocxReviewer) => {
  reviewer.acceptAll();
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
    { author: "Test" },
  );
  return reopened.snapshot().blocks.map((block) => ({
    text: block.text,
    inTable: block.table !== undefined,
  }));
};

/** Delete `texts` as tracked changes another author left pending. */
const withPendingDeletions = async (markdown: string, texts: string[]) => {
  const reviewer = await open(markdown);
  const result = apply(
    reviewer,
    "tracked-changes",
    texts.map((text, index) => ({
      id: `pending-${index}`,
      type: "deleteBlock",
      blockId: idOf(reviewer, text),
    })),
  );
  expect(result.skipped).toEqual([]);
  return FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
    author: "Test",
  });
};

const merge = (blockId: string): FolioDocumentOperation => ({
  id: "merge",
  type: "mergeBlockWithNext",
  blockId,
  separator: " ",
  mergedParagraphProperties: {
    numbering: { kind: "newList", format: "numbered" },
  },
});

type MergeCase = {
  markdown: string;
  /** Paragraphs another author's pending change deletes. */
  pending: string[];
  /** A batch before the merge, in the mode of the merge or as a suggestion. */
  before?: (reviewer: FolioDocxReviewer) => FolioDocumentOperation[];
  suggestedBefore?: boolean;
};

const mergeCases: Record<string, MergeCase> = {
  "a table that was there": {
    markdown: ["Opening.", "Kept item.", "Gone item.", TABLE, "End."].join(
      "\n\n",
    ),
    pending: ["Gone item."],
  },
  "two deleted paragraphs, then a table": {
    markdown: ["Kept item.", "Gone item.", "Gone too.", TABLE, "End."].join(
      "\n\n",
    ),
    pending: ["Gone item.", "Gone too."],
  },
  "a table an earlier batch inserted": {
    markdown: ["Kept item.", "Gone item.", "End."].join("\n\n"),
    pending: ["Gone item."],
    before: (reviewer) => [
      {
        id: "table",
        type: "insertTable",
        blockId: idOf(reviewer, "Gone item."),
        rows: ROWS,
      },
    ],
  },
  "a table an earlier suggestion inserted": {
    markdown: ["Kept item.", "Gone item.", "End."].join("\n\n"),
    pending: ["Gone item."],
    suggestedBefore: true,
    before: (reviewer) => [
      {
        id: "table",
        type: "insertTable",
        blockId: idOf(reviewer, "Gone item."),
        rows: ROWS,
      },
    ],
  },
};

describe("a merge whose next paragraphs accepting removes, before a table", () => {
  for (const [name, setup] of Object.entries(mergeCases)) {
    for (const mode of ["direct", "tracked-changes"] as const) {
      test(`is refused, as right before a table (${name}, ${mode})`, async () => {
        const reviewer = await withPendingDeletions(
          setup.markdown,
          setup.pending,
        );
        if (setup.before) {
          const before = apply(
            reviewer,
            setup.suggestedBefore === true ? "suggested" : mode,
            setup.before(reviewer),
          );
          expect(before.skipped).toEqual([]);
        }
        const blocksBefore = reviewer.snapshot().blocks;
        const result = apply(reviewer, mode, [
          merge(idOf(reviewer, "Kept item.")),
        ]);
        expect(result.skipped).toEqual([
          { id: "merge", reason: "unsupportedBlock" },
        ]);
        expect(result.applied).toEqual([]);
        expect(reviewer.snapshot().blocks).toEqual(blocksBefore);
      });
    }
  }

  test("still joins the paragraph past a table accepting removes", async () => {
    const markdown = ["Kept item.", "Gone item.", TABLE, "End."].join("\n\n");
    const results = [];
    for (const mode of ["direct", "tracked-changes"] as const) {
      const reviewer = await withPendingDeletions(markdown, ["Gone item."]);
      const removed = apply(reviewer, "tracked-changes", [
        {
          id: "table",
          type: "deleteTable",
          blockId: idOf(reviewer, "Widget"),
        },
      ]);
      expect(removed.skipped).toEqual([]);
      const result = apply(reviewer, mode, [
        merge(idOf(reviewer, "Kept item.")),
      ]);
      expect(result.skipped).toEqual([]);
      results.push(await accepted(reviewer));
    }
    const [direct, tracked] = results;
    expect(direct?.some(({ text }) => text.startsWith("Kept item. "))).toBe(
      true,
    );
    expect(tracked).toEqual(direct);
  });
});

type Insertion = {
  name: string;
  operation: (reviewer: FolioDocxReviewer) => FolioDocumentOperation;
};

/** Blocks after the paragraph before the deleted table. */
const afterOpening: Insertion[] = [
  {
    name: "a paragraph after the paragraph before it",
    operation: (reviewer) => ({
      id: "after",
      type: "insertAfterBlock",
      blockId: idOf(reviewer, "Opening clause."),
      text: "Inserted after.",
    }),
  },
  {
    name: "a table after the paragraph before it",
    operation: (reviewer) => ({
      id: "after",
      type: "insertTable",
      blockId: idOf(reviewer, "Opening clause."),
      rows: ROWS,
    }),
  },
];

/** Blocks before the paragraph after the deleted table. */
const beforeClosing: Insertion[] = [
  {
    name: "a paragraph before the paragraph after it",
    operation: (reviewer) => ({
      id: "before",
      type: "insertBeforeBlock",
      blockId: idOf(reviewer, "Closing clause."),
      text: "Inserted before.",
    }),
  },
  {
    name: "a table before the paragraph after it",
    operation: (reviewer) => ({
      id: "before",
      type: "insertTable",
      blockId: idOf(reviewer, "Closing clause."),
      position: "before",
      rows: [["Second", "Table"]],
    }),
  },
];

describe("blocks put on either side of a table accepting removes", () => {
  const markdown = ["Opening clause.", TABLE, "Closing clause."].join("\n\n");
  const replay = async (
    mode: "direct" | "tracked-changes",
    operations: (reviewer: FolioDocxReviewer) => FolioDocumentOperation[],
  ) => {
    const reviewer = await open(markdown);
    const removed = apply(reviewer, mode, [
      { id: "table", type: "deleteTable", blockId: idOf(reviewer, "Widget") },
    ]);
    expect(removed.skipped).toEqual([]);
    const result = apply(reviewer, mode, operations(reviewer));
    expect(result.skipped).toEqual([]);
    return accepted(reviewer);
  };
  for (const after of afterOpening) {
    for (const before of beforeClosing) {
      // Two new tables would meet once the deleted one goes: another rule.
      if (after.name.includes("table") && before.name.includes("table"))
        continue;
      for (const order of ["after first", "before first"] as const) {
        test(`share its gap in input order (${after.name}, ${before.name}, ${order})`, async () => {
          const operations = (reviewer: FolioDocxReviewer) => {
            const pair = [
              after.operation(reviewer),
              before.operation(reviewer),
            ];
            return order === "after first" ? pair : pair.toReversed();
          };
          const direct = await replay("direct", operations);
          expect(direct.map(({ text }) => text)).not.toContain("Widget");
          expect(await replay("tracked-changes", operations)).toEqual(direct);
        });
      }
    }
  }
});
