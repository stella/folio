/**
 * An edit reads a document with pending changes as accepting leaves it, past
 * the blocks accepting removes whole: applied as tracked changes and accepted,
 * it reads as the same edit applied directly.
 *
 * - A merge into a paragraph accepting removes, right before a table, ends
 *   there: the merged paragraph stays apart from the table in every mode,
 *   rather than running into its first cell when applied directly.
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

type Mode = "direct" | "tracked-changes";

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
    numbering: block.statedNumbering?.kind,
  }));
};

/**
 * Delete `texts` as tracked changes another author left pending; the ids of
 * every block, read before, since a reader shows a deleted one blank.
 */
const withPendingDeletions = async (markdown: string, texts: string[]) => {
  const reviewer = await open(markdown);
  const ids = new Map(
    reviewer.snapshot().blocks.map((block) => [block.text, block.id]),
  );
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
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
    { author: "Test" },
  );
  const id = (text: string): string => {
    const found = ids.get(text);
    if (found === undefined) throw new Error(`No block read "${text}"`);
    return found;
  };
  return { reviewer: reopened, id };
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
  /** A batch before the merge, in the mode of the merge. */
  before?: (id: (text: string) => string) => FolioDocumentOperation[];
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
    before: (id) => [
      {
        id: "table",
        type: "insertTable",
        blockId: id("Gone item."),
        rows: ROWS,
      },
    ],
  },
};

describe("a merge into a paragraph accepting removes, before a table", () => {
  const replay = async (
    setup: MergeCase,
    mode: "direct" | "tracked-changes",
  ) => {
    const { reviewer, id } = await withPendingDeletions(
      setup.markdown,
      setup.pending,
    );
    if (setup.before) {
      const before = apply(reviewer, mode, setup.before(id));
      expect(before.skipped).toEqual([]);
    }
    const result = apply(reviewer, mode, [merge(id("Kept item."))]);
    expect(result.skipped).toEqual([]);
    return accepted(reviewer);
  };
  for (const [name, setup] of Object.entries(mergeCases)) {
    test(`stays apart from the table, directly as tracked (${name})`, async () => {
      const direct = await replay(setup, "direct");
      expect(direct).toContainEqual({
        text: "Kept item. ",
        inTable: false,
        numbering: "reference",
      });
      expect(await replay(setup, "tracked-changes")).toEqual(direct);
    });
  }
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
