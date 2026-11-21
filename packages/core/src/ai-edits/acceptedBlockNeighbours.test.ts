/**
 * An edit reads a document with pending changes as accepting leaves it, past
 * the blocks accepting removes whole: applied as tracked changes and accepted,
 * it reads as the same edit applied directly.
 *
 * - A merge into a paragraph accepting removes, right before a table: on the
 *   legacy session it ends there, the merged paragraph apart from the table;
 *   on the canonical session the merged words run into the table's first
 *   cell. Either way, directly as tracked.
 * - Blocks put on either side of blocks accepting removes whole, tables and
 *   paragraphs, share one gap and keep the order the batch names them in, as
 *   they do once those blocks are removed directly.
 * - Removing what ends a pending join leaves the joined paragraph standing:
 *   deleting the paragraph the join runs into keeps its break, the joined
 *   paragraph's, and on the legacy session removing a table that stops the
 *   join keeps the words apart from what follows it, as accepting already
 *   left them.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

type Mode = "direct" | "tracked-changes";

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const TABLE = "| Item | Price |\n| --- | --- |\n| Widget | 10 |";
const ROWS = [
  ["Term", "Value"],
  ["Period", "12 months"],
];

/** Opens `bytes` on `session`, or on the reviewer's default without one. */
const reviewerOf = (bytes: ArrayBuffer, session?: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(bytes, {
    author: "Test",
    ...(session !== undefined && { session }),
  });

const open = async (
  markdown: string,
  session?: HeadlessReviewerSession,
): Promise<FolioDocxReviewer> =>
  reviewerOf(
    (await ensureParaIds(await createDocx(fromMarkdown(markdown)))).docx,
    session,
  );

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
const accepted = async (
  reviewer: FolioDocxReviewer,
  session?: HeadlessReviewerSession,
) => {
  reviewer.acceptAll();
  const reopened = await reviewerOf(await reviewer.toBuffer(), session);
  return reopened.snapshot().blocks.map((block) => ({
    text: block.text,
    inTable: block.table !== undefined,
    numbering: block.statedNumbering?.kind,
  }));
};

type PendingDeletionsOptions = {
  markdown: string;
  /** The texts of the paragraphs deleted. */
  texts: string[];
  session?: HeadlessReviewerSession;
};

/**
 * Delete `texts` as tracked changes another author left pending; the ids of
 * every block, read before, since a reader shows a deleted one blank.
 */
const withPendingDeletions = async ({
  markdown,
  texts,
  session,
}: PendingDeletionsOptions) => {
  const reviewer = await open(markdown, session);
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
  const reopened = await reviewerOf(await reviewer.toBuffer(), session);
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
  /** The table's first cell. */
  cell: string;
  /** A batch before the merge, in the mode of the merge. */
  before?: (id: (text: string) => string) => FolioDocumentOperation[];
};

const mergeCases: Record<string, MergeCase> = {
  "a table that was there": {
    markdown: ["Opening.", "Kept item.", "Gone item.", TABLE, "End."].join(
      "\n\n",
    ),
    pending: ["Gone item."],
    cell: "Item",
  },
  "two deleted paragraphs, then a table": {
    markdown: ["Kept item.", "Gone item.", "Gone too.", TABLE, "End."].join(
      "\n\n",
    ),
    pending: ["Gone item.", "Gone too."],
    cell: "Item",
  },
  "a table an earlier batch inserted": {
    markdown: ["Kept item.", "Gone item.", "End."].join("\n\n"),
    pending: ["Gone item."],
    cell: "Term",
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
    session: HeadlessReviewerSession,
  ) => {
    const { reviewer, id } = await withPendingDeletions({
      markdown: setup.markdown,
      texts: setup.pending,
      session,
    });
    if (setup.before) {
      const before = apply(reviewer, mode, setup.before(id));
      expect(before.skipped).toEqual([]);
    }
    const result = apply(reviewer, mode, [merge(id("Kept item."))]);
    return {
      skipped: result.skipped,
      accepted: await accepted(reviewer, session),
    };
  };
  for (const [name, setup] of Object.entries(mergeCases)) {
    test(`stays apart from the table, directly as tracked (${name}, legacy)`, async () => {
      const direct = await replay(setup, "direct", "legacy");
      expect(direct.skipped).toEqual([]);
      expect(direct.accepted).toContainEqual({
        text: "Kept item. ",
        inTable: false,
        numbering: "reference",
      });
      expect(await replay(setup, "tracked-changes", "legacy")).toEqual(direct);
    });

    // The merged words take the cell paragraph's mark, and its properties.
    test(`runs into the table's first cell, directly as tracked (${name}, canonical)`, async () => {
      const direct = await replay(setup, "direct", "canonical");
      expect(direct.skipped).toEqual([]);
      expect(direct.accepted).toContainEqual({
        text: `Kept item. ${setup.cell}`,
        inTable: true,
        numbering: "inherit",
      });
      const tracked = await replay(setup, "tracked-changes", "canonical");
      if (setup.before === undefined) {
        expect(tracked).toEqual(direct);
        return;
      }
      // A join through a deleted paragraph into a pending inserted table
      // is refused.
      expect(tracked.skipped).toEqual([
        expect.objectContaining({
          id: "merge",
          message: expect.stringContaining("would run it into the table"),
        }),
      ]);
      expect(tracked.accepted.at(0)?.text).toBe("Kept item.");
    });
  }
});

type Insertion = {
  name: string;
  operation: (reviewer: FolioDocxReviewer) => FolioDocumentOperation;
  /** Text the first block it puts in holds. */
  marker: string;
  /** Applied directly only: it has no tracked form. */
  directOnly?: true;
};

/** Blocks after the paragraph before the deleted table. */
const afterOpening: Insertion[] = [
  {
    name: "a paragraph after the paragraph before it",
    marker: "Inserted after.",
    operation: (reviewer) => ({
      id: "after",
      type: "insertAfterBlock",
      blockId: idOf(reviewer, "Opening clause."),
      text: "Inserted after.",
    }),
  },
  {
    name: "a table after the paragraph before it",
    marker: "Term",
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
    marker: "Inserted before.",
    operation: (reviewer) => ({
      id: "before",
      type: "insertBeforeBlock",
      blockId: idOf(reviewer, "Closing clause."),
      text: "Inserted before.",
    }),
  },
  {
    name: "a table before the paragraph after it",
    marker: "Second",
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
  for (const session of SESSIONS) {
    const replay = async (
      mode: "direct" | "tracked-changes",
      operations: (reviewer: FolioDocxReviewer) => FolioDocumentOperation[],
    ) => {
      const reviewer = await open(markdown, session);
      const removed = apply(reviewer, mode, [
        { id: "table", type: "deleteTable", blockId: idOf(reviewer, "Widget") },
      ]);
      expect(removed.skipped).toEqual([]);
      const result = apply(reviewer, mode, operations(reviewer));
      expect(result.skipped).toEqual([]);
      return accepted(reviewer, session);
    };
    for (const after of afterOpening) {
      for (const before of beforeClosing) {
        // Two new tables would meet once the deleted one goes: another rule.
        if (after.name.includes("table") && before.name.includes("table"))
          continue;
        for (const order of ["after first", "before first"] as const) {
          test(`share its gap in input order (${after.name}, ${before.name}, ${order}, ${session})`, async () => {
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
  }
});

/** The block texts once every change is rejected, from a saved copy. */
const rejected = async (
  reviewer: FolioDocxReviewer,
  session?: HeadlessReviewerSession,
): Promise<string[]> => {
  const copy = await reviewerOf(await reviewer.toBuffer(), session);
  copy.rejectAll();
  return copy.snapshot().blocks.map((block) => block.text);
};

const textsOf = (blocks: readonly { text: string }[]): string[] =>
  blocks.map(({ text }) => text);

const joinInto = (blockId: string): FolioDocumentOperation => ({
  id: "join",
  type: "mergeBlockWithNext",
  blockId,
  separator: " ",
});

const deleteBlock = (
  blockId: string,
  id = "delete",
): FolioDocumentOperation => ({
  id,
  type: "deleteBlock",
  blockId,
});

type IdOf = (text: string) => string;
type Step = (id: IdOf) => FolioDocumentOperation[];

const deleteTable: Step = (id) => [
  { id: "table", type: "deleteTable", blockId: id("Widget") },
];

const tableRemovals: Record<string, Step> = {
  deleteTable,
  deleteTableRow: (id) => [
    { id: "header", type: "deleteTableRow", blockId: id("Item") },
    { id: "row", type: "deleteTableRow", blockId: id("Widget") },
  ],
};

describe("removing what ends a pending join", () => {
  /** Each step a tracked batch; the texts once accepted. */
  const replay = async (
    markdown: string,
    steps: Step[],
    session?: HeadlessReviewerSession,
  ) => {
    const reviewer = await open(markdown, session);
    const before = textsOf(reviewer.snapshot().blocks);
    const ids = new Map(before.map((text, index) => [text, index]));
    const blocks = reviewer.snapshot().blocks;
    const id: IdOf = (text) => {
      const index = ids.get(text);
      const block = index === undefined ? undefined : blocks[index];
      if (!block) throw new Error(`No block read "${text}"`);
      return block.id;
    };
    for (const step of steps) {
      expect(apply(reviewer, "tracked-changes", step(id)).skipped).toEqual([]);
    }
    // Rejecting every change still gives the document before back.
    expect(await rejected(reviewer, session)).toEqual(before);
    return textsOf(await accepted(reviewer, session));
  };

  const joined: Step[] = [
    (id) => [joinInto(id("Opening."))],
    (id) => [deleteBlock(id("Joined."))],
  ];

  test("keeps the break of the paragraph the join runs into", async () => {
    expect(
      await replay(["Opening.", "Joined.", "Closing."].join("\n\n"), joined),
    ).toEqual(["Opening. ", "Closing."]);
  });

  // The rest of a batch reads the joined paragraph where its deletion left
  // it, as after that deletion alone: a list item the join runs into.
  const beforeListItem = ["- Opening clause.", "- Joined.", "Closing."].join(
    "\n\n",
  );
  type AfterDeletionOptions = {
    session: HeadlessReviewerSession;
    edit: (id: IdOf) => FolioDocumentOperation;
    together: boolean;
  };
  const afterDeletion = async ({
    session,
    edit,
    together,
  }: AfterDeletionOptions) => {
    const reviewer = await open(beforeListItem, session);
    const id: IdOf = (text) => idOf(reviewer, text);
    const opening = id("Opening clause.");
    const joinedId = id("Joined.");
    expect(
      apply(reviewer, "tracked-changes", [joinInto(opening)]).skipped,
    ).toEqual([]);
    const operations = [edit(() => opening), deleteBlock(joinedId)];
    const skipped = together
      ? apply(reviewer, "tracked-changes", operations).skipped
      : [
          ...apply(reviewer, "tracked-changes", [operations[1]!]).skipped,
          ...apply(reviewer, "tracked-changes", [operations[0]!]).skipped,
        ];
    return {
      skipped: skipped.map(({ reason }) => reason),
      blocks: reviewer.snapshot().blocks.map(({ text, kind }) => ({
        text,
        kind,
      })),
    };
  };

  const editsOfTheSource: Record<string, (id: IdOf) => FolioDocumentOperation> =
    {
      "a split keeps the list item's properties": (id) => ({
        id: "split",
        type: "splitBlock",
        blockId: id("Opening clause."),
        offset: 8,
      }),
      "a merge meets the break the deletion kept": (id) => ({
        ...joinInto(id("Opening clause.")),
        id: "merge",
        mergedParagraphProperties: { alignment: "center" },
      }),
    };

  for (const session of SESSIONS) {
    for (const [name, edit] of Object.entries(editsOfTheSource)) {
      test(`reads it in the same batch as after the deletion alone (${name}, ${session})`, async () => {
        expect(await afterDeletion({ session, edit, together: true })).toEqual(
          await afterDeletion({ session, edit, together: false }),
        );
      });
    }

    test(`lets that break go with the words joined into it (${session})`, async () => {
      expect(
        await replay(
          ["Opening.", "Joined.", "Closing."].join("\n\n"),
          [
            (id) => [joinInto(id("Opening."))],
            (id) => [
              deleteBlock(id("Opening."), "opening"),
              deleteBlock(id("Joined."), "joined"),
            ],
          ],
          session,
        ),
      ).toEqual(["Closing."]);
    });

    test(`lets that break go when a later batch deletes the joined words (${session})`, async () => {
      expect(
        await replay(
          ["Opening.", "Joined.", "Closing."].join("\n\n"),
          [
            (id) => [joinInto(id("Opening."))],
            (id) => [deleteBlock(id("Joined."), "joined")],
            (id) => [deleteBlock(id("Opening."), "opening")],
          ],
          session,
        ),
      ).toEqual(["Closing."]);
    });
  }

  const beforeTable = ["Opening.", "Joined.", TABLE, "Closing."].join("\n\n");

  test("keeps it before a table", async () => {
    expect(await replay(beforeTable, joined)).toEqual([
      "Opening. ",
      "Item",
      "Price",
      "Widget",
      "10",
      "Closing.",
    ]);
  });

  for (const [name, removal] of Object.entries(tableRemovals)) {
    test(`keeps it once that table goes (${name})`, async () => {
      expect(await replay(beforeTable, [...joined, removal])).toEqual([
        "Opening. ",
        "Closing.",
      ]);
    });

    for (const session of SESSIONS)
      test(`gives the words the join reached once the table goes, directly as tracked (${name}, ${session})`, async () => {
        const run = async (mode: Mode, step: Step) => {
          const { reviewer, id } = await withPendingDeletions({
            markdown: ["Kept item.", "Gone item.", TABLE, "End."].join("\n\n"),
            texts: ["Gone item."],
            session,
          });
          expect(
            apply(reviewer, mode, [merge(id("Kept item."))]).skipped,
          ).toEqual([]);
          expect(apply(reviewer, mode, step(id)).skipped).toEqual([]);
          return accepted(reviewer, session);
        };
        const direct = await run("direct", deleteTable);
        // The legacy session keeps the words a table stopped apart; on the
        // canonical session the join runs on to what follows the table.
        expect(textsOf(direct)).toEqual(
          session === "legacy" ? ["Kept item. ", "End."] : ["Kept item. End."],
        );
        expect(await run("tracked-changes", removal)).toEqual(direct);
      });
  }
});

const signatureTable = (
  id: string,
  blockId: string,
  position: "after" | "before",
): FolioDocumentOperation => ({
  id,
  type: "insertSignatureTable",
  blockId,
  position,
  parties: [{ name: "Signer" }],
});

/** Every block insertion after the paragraph before the removed blocks. */
const everyAfterOpening: Insertion[] = [
  ...afterOpening,
  {
    name: "a signature table after the paragraph before it",
    marker: "Signer",
    operation: (reviewer) =>
      signatureTable("after", idOf(reviewer, "Opening clause."), "after"),
    directOnly: true,
  },
];

/** Every block insertion before the paragraph after the removed blocks. */
const everyBeforeClosing: Insertion[] = [
  ...beforeClosing,
  {
    name: "a signature table before the paragraph after it",
    marker: "Signer",
    operation: (reviewer) =>
      signatureTable("before", idOf(reviewer, "Closing clause."), "before"),
    directOnly: true,
  },
];

type RemovedRun = {
  /** The blocks between the two anchors. */
  between: string[];
  /** Removes them whole, by the ids read before. */
  removal: (id: IdOf) => FolioDocumentOperation[];
};

const removeParagraph =
  (text: string, id = text): Step =>
  (idOfText) => [deleteBlock(idOfText(text), id)];

const removedRuns: Record<string, RemovedRun> = {
  "a paragraph": {
    between: ["Removed clause."],
    removal: removeParagraph("Removed clause."),
  },
  "two paragraphs": {
    between: ["Removed clause.", "Removed too."],
    removal: (id) => [
      ...removeParagraph("Removed clause.")(id),
      ...removeParagraph("Removed too.")(id),
    ],
  },
  "a paragraph, then a table": {
    between: ["Removed clause.", TABLE],
    removal: (id) => [
      ...removeParagraph("Removed clause.")(id),
      ...deleteTable(id),
    ],
  },
  "a table, then a paragraph": {
    between: [TABLE, "Removed clause."],
    removal: (id) => [
      ...deleteTable(id),
      ...removeParagraph("Removed clause.")(id),
    ],
  },
  "a paragraph, a table and a paragraph": {
    between: ["Removed clause.", TABLE, "Removed too."],
    removal: (id) => [
      ...removeParagraph("Removed clause.")(id),
      ...deleteTable(id),
      ...removeParagraph("Removed too.")(id),
    ],
  },
};

describe("blocks put on either side of blocks accepting removes whole", () => {
  /**
   * The removal in `removalMode`, then the batch in `mode`: the texts the
   * source had, and the blocks once every change is accepted.
   */
  type ReplayOptions = {
    session: HeadlessReviewerSession;
    run: RemovedRun;
    removalMode: Mode;
    mode: Mode;
    operations: (reviewer: FolioDocxReviewer) => FolioDocumentOperation[];
  };
  const replay = async ({
    session,
    run,
    removalMode,
    mode,
    operations,
  }: ReplayOptions) => {
    const reviewer = await open(
      ["Opening clause.", ...run.between, "Closing clause."].join("\n\n"),
      session,
    );
    const source = textsOf(reviewer.snapshot().blocks);
    const ids = new Map(
      reviewer.snapshot().blocks.map((block) => [block.text, block.id]),
    );
    const id: IdOf = (text) => {
      const found = ids.get(text);
      if (found === undefined) throw new Error(`No block read "${text}"`);
      return found;
    };
    expect(apply(reviewer, removalMode, run.removal(id)).skipped).toEqual([]);
    expect(apply(reviewer, mode, operations(reviewer)).skipped).toEqual([]);
    if (removalMode === "tracked-changes" && mode === "tracked-changes") {
      // Rejecting every change still gives the source back.
      expect(await rejected(reviewer, session)).toEqual(source);
    }
    return accepted(reviewer, session);
  };
  for (const session of SESSIONS)
    for (const [runName, run] of Object.entries(removedRuns)) {
      for (const after of everyAfterOpening) {
        for (const before of everyBeforeClosing) {
          // Two new tables would meet once the removed blocks go: another rule.
          if (after.name.includes("table") && before.name.includes("table"))
            continue;
          for (const order of ["after first", "before first"] as const) {
            test(`share their gap in input order (${runName}; ${after.name}, ${before.name}, ${order}, ${session})`, async () => {
              const operations = (reviewer: FolioDocxReviewer) => {
                const pair = [
                  after.operation(reviewer),
                  before.operation(reviewer),
                ];
                return order === "after first" ? pair : pair.toReversed();
              };
              const options = { session, run, operations };
              const direct = await replay({
                ...options,
                removalMode: "direct",
                mode: "direct",
              });
              const texts = textsOf(direct);
              expect(texts).not.toContain("Removed clause.");
              // Between the two paragraphs, in the order the batch names them.
              const indexOf = (marker: string): number =>
                texts.findIndex((text) => text.includes(marker));
              const [first, second] =
                order === "after first" ? [after, before] : [before, after];
              expect(indexOf("Opening clause.")).toBeLessThan(
                indexOf(first.marker),
              );
              expect(indexOf(first.marker)).toBeLessThan(
                indexOf(second.marker),
              );
              expect(indexOf(second.marker)).toBeLessThan(
                indexOf("Closing clause."),
              );
              if (!after.directOnly && !before.directOnly) {
                expect(
                  await replay({
                    ...options,
                    removalMode: "tracked-changes",
                    mode: "tracked-changes",
                  }),
                ).toEqual(direct);
              }
              expect(
                await replay({
                  ...options,
                  removalMode: "tracked-changes",
                  mode: "direct",
                }),
              ).toEqual(direct);
            });
          }
        }
      }
    }

  for (const before of beforeClosing) {
    test(`stop at a deleted paragraph whose break a pending join keeps (${before.name})`, async () => {
      const replayJoin = async (mode: Mode) => {
        const reviewer = await open(
          ["Opening clause.", "Joined clause.", "Closing clause."].join("\n\n"),
        );
        const source = textsOf(reviewer.snapshot().blocks);
        const joined = idOf(reviewer, "Joined clause.");
        expect(
          apply(reviewer, "tracked-changes", [
            joinInto(idOf(reviewer, "Opening clause.")),
          ]).skipped,
        ).toEqual([]);
        expect(
          apply(reviewer, "tracked-changes", [deleteBlock(joined)]).skipped,
        ).toEqual([]);
        expect(
          apply(reviewer, mode, [before.operation(reviewer)]).skipped,
        ).toEqual([]);
        if (mode === "tracked-changes") {
          expect(await rejected(reviewer)).toEqual(source);
        }
        return accepted(reviewer);
      };
      const direct = await replayJoin("direct");
      // The joined words keep that break, apart from the new block.
      expect(direct[0]).toMatchObject({
        text: "Opening clause. ",
        inTable: false,
      });
      expect(direct.at(-1)?.text).toBe("Closing clause.");
      expect(await replayJoin("tracked-changes")).toEqual(direct);
    });
  }

  const reachedTable = async (session: HeadlessReviewerSession, mode: Mode) => {
    const { reviewer, id } = await withPendingDeletions({
      markdown: ["Kept item.", "Gone item.", TABLE, "End."].join("\n\n"),
      texts: ["Gone item."],
      session,
    });
    expect(apply(reviewer, mode, [merge(id("Kept item."))]).skipped).toEqual(
      [],
    );
    return { reviewer, id };
  };

  // The join runs into the table's first cell: a table put before it would
  // stand inside the joined paragraph.
  for (const mode of ["direct", "tracked-changes"] as const)
    test(`before a table a pending join runs into, is refused (${mode}, canonical)`, async () => {
      const { reviewer, id } = await reachedTable("canonical", mode);
      const before = reviewer.snapshot();
      expect(
        apply(reviewer, mode, [
          {
            id: "table",
            type: "insertTable",
            blockId: id("Widget"),
            position: "before",
            rows: ROWS,
          },
        ]).skipped.map(({ reason }) => reason),
      ).toEqual(["pendingParagraphMarkDeletion"]);
      expect(reviewer.snapshot()).toEqual(before);
    });

  test("stop where a pending join reaches a paragraph accepting removes (legacy)", async () => {
    const session = "legacy";
    const replayReached = async (mode: Mode) => {
      const { reviewer, id } = await withPendingDeletions({
        markdown: ["Kept item.", "Gone item.", TABLE, "End."].join("\n\n"),
        texts: ["Gone item."],
        session,
      });
      expect(apply(reviewer, mode, [merge(id("Kept item."))]).skipped).toEqual(
        [],
      );
      expect(
        apply(reviewer, mode, [
          {
            id: "table",
            type: "insertTable",
            blockId: id("Widget"),
            position: "before",
            rows: ROWS,
          },
        ]).skipped,
      ).toEqual([]);
      return accepted(reviewer, session);
    };
    const direct = await replayReached("direct");
    expect(direct[0]).toEqual({
      text: "Kept item. ",
      inTable: false,
      numbering: "reference",
    });
    expect(textsOf(direct).slice(1, 5)).toEqual([
      "Term",
      "Value",
      "Period",
      "12 months",
    ]);
    expect(await replayReached("tracked-changes")).toEqual(direct);
  });
});

describe("a paragraph put before a block the batch restyles, past blocks accepting removes whole", () => {
  // The batch orders the insertion as at the restyled block, as it does with
  // nothing between them: the paragraph takes what the block had.
  for (const session of SESSIONS) {
    const inserted = async (pending: string[]) => {
      const { reviewer, id } = await withPendingDeletions({
        markdown: ["Opening clause.", "Gone clause.", "Closing clause."].join(
          "\n\n",
        ),
        texts: pending,
        session,
      });
      const result = apply(reviewer, "tracked-changes", [
        {
          id: "restyle",
          type: "setBlockParagraphProperties",
          blockId: id("Closing clause."),
          properties: { styleId: "Heading2" },
        },
        {
          id: "insert",
          type: "insertBeforeBlock",
          blockId: id("Closing clause."),
          text: "Inserted clause.",
        },
      ]);
      expect(result.skipped).toEqual([]);
      const block = reviewer
        .snapshot()
        .blocks.find(({ text }) => text === "Inserted clause.");
      return { kind: block?.kind, styleId: block?.styleId };
    };

    test(`takes the same properties as with nothing between them (${session})`, async () => {
      const plain = await inserted([]);
      expect(plain).toEqual({ kind: "paragraph", styleId: undefined });
      expect(await inserted(["Gone clause."])).toEqual(plain);
    });
  }
});

// The canonical session puts it after the final paragraph, whose mark holds
// the pending formatting change the deletions recorded: moving that change
// onto a new final paragraph is refused.
const FINAL_MARK_REFUSAL =
  "A tracked insertion after a paragraph no paragraph follows needs the anchor's mark free of review records.";

describe("a paragraph put right after words whose break a pending change deletes", () => {
  // The words run, accepted, into the closing paragraph; a paragraph put
  // after them takes that deleted break, and the words keep their own. The
  // rest of the batch reads them so, as after that insertion alone.
  type AfterInsertionOptions = {
    session: HeadlessReviewerSession;
    together: boolean;
  };
  const afterInsertion = async ({
    session,
    together,
  }: AfterInsertionOptions) => {
    const { reviewer, id } = await withPendingDeletions({
      markdown: [
        "- Opening clause.",
        "Middle clause.",
        "- Closing clause.",
      ].join("\n\n"),
      texts: ["Middle clause.", "Closing clause."],
      session,
    });
    const operations: FolioDocumentOperation[] = [
      {
        id: "insert",
        type: "insertAfterBlock",
        blockId: id("Opening clause."),
        text: "Inserted clause.",
        styleId: "Heading2",
      },
      {
        id: "split",
        type: "splitBlock",
        blockId: id("Opening clause."),
        offset: 8,
      },
    ];
    const skipped = together
      ? apply(reviewer, "tracked-changes", operations).skipped
      : operations.flatMap(
          (operation) =>
            apply(reviewer, "tracked-changes", [operation]).skipped,
        );
    expect(skipped).toEqual(
      session === "canonical"
        ? [
            expect.objectContaining({
              id: "insert",
              message: FINAL_MARK_REFUSAL,
            }),
          ]
        : [],
    );
    return reviewer
      .snapshot()
      .blocks.map(({ text, kind, styleId }) => ({ text, kind, styleId }));
  };

  for (const session of SESSIONS)
    test(`the rest of the batch reads the words with their own properties (${session})`, async () => {
      const alone = await afterInsertion({ session, together: false });
      expect(alone.find(({ text }) => text === "Opening ")).toEqual({
        text: "Opening ",
        kind: "listItem",
        styleId: undefined,
      });
      expect(await afterInsertion({ session, together: true })).toEqual(alone);
    });
});
