/**
 * A block inserted beside a paragraph whose break a tracked merge removes
 * treats the two as the joined paragraph they become: a paragraph copies the
 * properties the joined one will have, and a block after it goes after the
 * paragraph the join leaves. Accepting gives what the same edits applied
 * directly give.
 *
 * The canonical session puts a paragraph after one whose break is pending
 * deletion right after it, with its properties, while the paragraph the join
 * leaves keeps words once accepted: accepting runs its words into the new
 * paragraph and leaves the other one standing (`insertionAfterPendingBreak`).
 */

import { describe, expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const MARKDOWN = "# Terms\n\n1. Numbered clause one\n\nUnnumbered body text.";

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

/** Opens `bytes` on `session`, or on the reviewer's default without one. */
const reviewerOf = (bytes: ArrayBuffer, session?: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(bytes, {
    author: "Test",
    ...(session !== undefined && { session }),
  });

const openMarkdown = async (
  markdown: string,
  session?: HeadlessReviewerSession,
) =>
  reviewerOf(
    (await ensureParaIds(await createDocx(fromMarkdown(markdown)))).docx,
    session,
  );

const open = (session?: HeadlessReviewerSession) =>
  openMarkdown(MARKDOWN, session);

type MergedProperties = Extract<
  FolioDocumentOperation,
  { type: "mergeBlockWithNext" }
>["mergedParagraphProperties"];

const requests: Record<string, (numId: number) => MergedProperties> = {
  "the next paragraph's list": (numId) => ({
    numbering: { kind: "reference", numId, ilvl: 0 },
  }),
  "another style": () => ({ styleId: "Heading2" }),
  "no list": () => ({ numbering: { kind: "none" } }),
  "nothing more": () => undefined,
};

const insertions = {
  "a paragraph before it": (blockId: string): FolioDocumentOperation => ({
    id: "insert",
    type: "insertBeforeBlock",
    blockId,
    text: "Inserted clause.",
  }),
  "a paragraph after it": (blockId: string): FolioDocumentOperation => ({
    id: "insert",
    type: "insertAfterBlock",
    blockId,
    text: "Inserted clause.",
  }),
  "a table after it": (blockId: string): FolioDocumentOperation => ({
    id: "insert",
    type: "insertTable",
    blockId,
    rows: [
      ["Term", "Value"],
      ["Period", "12 months"],
    ],
  }),
};

type EditOptions = {
  session: HeadlessReviewerSession;
  mode: "direct" | "tracked-changes";
  batched: boolean;
  merged: (numId: number) => MergedProperties;
  insertion: (blockId: string) => FolioDocumentOperation;
};

type BlockShape = Pick<
  ReturnType<FolioDocxReviewer["snapshot"]>["blocks"][number],
  "text" | "styleId" | "statedNumbering" | "listReference" | "table"
>;

const shapeOf = ({
  text,
  styleId,
  statedNumbering,
  listReference,
  table,
}: BlockShape): BlockShape => ({
  text,
  styleId,
  statedNumbering,
  listReference,
  table,
});

/** Merge the heading into the list item, insert beside it, accept, and read. */
const edit = async ({
  session,
  mode,
  batched,
  merged,
  insertion,
}: EditOptions) => {
  const reviewer = await open(session);
  const [heading, item] = reviewer.snapshot().blocks;
  const numId = item?.listReference?.numId;
  if (!heading || numId === undefined) throw new Error("Missing fixture");
  const properties = merged(numId);
  const merge: FolioDocumentOperation = {
    id: "merge",
    type: "mergeBlockWithNext",
    blockId: heading.id,
    separator: " ",
    ...(properties !== undefined && { mergedParagraphProperties: properties }),
  };
  const applies = (operations: FolioDocumentOperation[]) => {
    const { applied, skipped } = reviewer.applyOperations(operations, {
      mode,
      atomic: true,
    });
    expect(skipped).toEqual([]);
    return applied;
  };
  if (batched) applies([merge, insertion(heading.id)]);
  else {
    // A later batch names the joined paragraph as it now stands: applied
    // directly, the join keeps the later paragraph's id.
    const joined = applies([merge]).find(
      (entry) => entry.paragraphIdentity !== undefined,
    )?.paragraphIdentity;
    applies([
      insertion(
        joined?.type === "joined" && joined.retiredBlockId === heading.id
          ? joined.blockId
          : heading.id,
      ),
    ]);
  }
  if (mode !== "direct") reviewer.acceptAll();
  const reopened = await reviewerOf(await reviewer.toBuffer(), session);
  return reopened.snapshot().blocks.map(shapeOf);
};

describe("an insertion beside a paragraph a tracked merge joins", () => {
  for (const session of SESSIONS) {
    for (const [placed, insertion] of Object.entries(insertions)) {
      for (const [name, merged] of Object.entries(requests)) {
        // One batch cannot both join a paragraph and place a block at its end.
        for (const batched of placed.endsWith("before it")
          ? [true, false]
          : [false]) {
          test(`reads as the joined paragraph (${placed}, ${name}, ${batched ? "one batch" : "two batches"}, ${session})`, async () => {
            const options = { session, batched, merged, insertion };
            const direct = await edit({ ...options, mode: "direct" });
            expect(
              direct.some(
                ({ text }) => text.startsWith("Inserted") || text === "Term",
              ),
            ).toBe(true);
            const tracked = await edit({ ...options, mode: "tracked-changes" });
            if (session === "legacy" || placed !== "a paragraph after it") {
              expect(tracked).toEqual(direct);
              return;
            }
            // The heading's words run into the new paragraph, which has the
            // heading's properties; the list item stays, with the properties
            // the merge gave the joined paragraph.
            const [joined, , ...rest] = direct;
            const heading = (await open(session)).snapshot().blocks.at(0);
            if (!joined || !heading) throw new Error("Missing fixture");
            expect(tracked).toEqual([
              { ...shapeOf(heading), text: "Terms Inserted clause." },
              { ...joined, text: "Numbered clause one" },
              ...rest,
            ]);
          });
        }
      }
    }
  }
});

type AroundOptions = {
  /** Which of the two insertions comes first. */
  order: "after first" | "before first";
  batched: boolean;
};

/**
 * After a tracked merge, insert after the joined heading and before the
 * paragraph it joins, accept, and read the texts.
 */
const around = async ({ order, batched }: AroundOptions) => {
  const reviewer = await open();
  const [heading, item] = reviewer.snapshot().blocks;
  if (!heading || !item) throw new Error("Missing fixture");
  reviewer.applyOperations(
    [
      {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: heading.id,
        separator: " ",
      },
    ],
    { mode: "tracked-changes" },
  );
  const after: FolioDocumentOperation = {
    id: "after",
    type: "insertAfterBlock",
    blockId: heading.id,
    text: "After.",
  };
  const before: FolioDocumentOperation = {
    id: "before",
    type: "insertBeforeBlock",
    blockId: item.id,
    text: "Before.",
  };
  const operations =
    order === "after first" ? [after, before] : [before, after];
  const batches = batched
    ? [operations]
    : operations.map((operation) => [operation]);
  for (const batch of batches) {
    reviewer.applyOperations(batch, { mode: "tracked-changes" });
  }
  reviewer.acceptAll();
  return reviewer.snapshot().blocks.map(({ text }) => text);
};

describe("insertions on both sides of a pending join", () => {
  for (const order of ["after first", "before first"] as const) {
    test(`leave the joined words together, in a batch as one at a time (${order})`, async () => {
      const sequential = await around({ order, batched: false });
      expect(await around({ order, batched: true })).toEqual(sequential);
      const joined = sequential.indexOf("Terms Numbered clause one");
      expect(joined).toBeGreaterThan(-1);
      expect(sequential.indexOf("Before.")).toBe(joined - 1);
      expect(sequential.indexOf("After.")).toBe(joined + 1);
    });
  }
  for (const session of SESSIONS)
    test(`after a join through a paragraph deleted whole goes where the direct edits put it (${session})`, async () => {
      const read = async (mode: "direct" | "tracked-changes") => {
        const reviewer = await open(session);
        const [heading, item] = reviewer.snapshot().blocks;
        if (!heading || !item) throw new Error("Missing fixture");
        const joinedIds = new Map<string, string>();
        for (const operation of [
          { id: "delete", type: "deleteBlock", blockId: item.id },
          {
            id: "merge",
            type: "mergeBlockWithNext",
            blockId: heading.id,
            separator: " ",
          },
          {
            id: "after",
            type: "insertAfterBlock",
            blockId: heading.id,
            text: "After.",
          },
        ] satisfies FolioDocumentOperation[]) {
          const result = reviewer.applyOperations(
            [
              {
                ...operation,
                blockId: joinedIds.get(operation.blockId) ?? operation.blockId,
              },
            ],
            { mode },
          );
          expect(result.skipped).toEqual([]);
          for (const { paragraphIdentity: joined } of result.applied) {
            if (joined?.type === "joined")
              joinedIds.set(joined.retiredBlockId, joined.blockId);
          }
        }
        if (mode !== "direct") reviewer.acceptAll();
        return reviewer.snapshot().blocks.map(({ text }) => text);
      };
      const direct = await read("direct");
      expect(direct).toEqual(["Terms Unnumbered body text.", "After."]);
      // On the canonical session the paragraph the join leaves keeps words, so
      // the new one goes right after the heading and takes its words.
      expect(await read("tracked-changes")).toEqual(
        session === "canonical"
          ? ["Terms After.", "Unnumbered body text."]
          : direct,
      );
    });
});

// The canonical session puts it after the final paragraph, whose mark holds
// the pending formatting change the deletion recorded: moving that change
// onto a new final paragraph is refused.
const FINAL_MARK_REFUSAL =
  "A tracked insertion after a paragraph no paragraph follows needs the anchor's mark free of review records.";

describe("an insertion after a paragraph whose break a final deletion retires", () => {
  for (const session of SESSIONS)
    test(`stays beside it: the emptied final paragraph has no words to join (${session})`, async () => {
      const reviewer = await open(session);
      const blocks = reviewer.snapshot().blocks;
      const [item, last] = [blocks.at(1), blocks.at(-1)];
      if (!item || !last) throw new Error("Missing fixture");
      reviewer.applyOperations(
        [{ id: "delete", type: "deleteBlock", blockId: last.id }],
        {
          mode: "tracked-changes",
        },
      );
      const before = reviewer.snapshot();
      const { skipped } = reviewer.applyOperations(
        [
          {
            id: "after",
            type: "insertAfterBlock",
            blockId: item.id,
            text: "After.",
          },
        ],
        { mode: "tracked-changes" },
      );
      if (session === "canonical") {
        expect(skipped).toEqual([
          expect.objectContaining({ id: "after", message: FINAL_MARK_REFUSAL }),
        ]);
        expect(reviewer.snapshot()).toEqual(before);
        return;
      }
      expect(skipped).toEqual([]);
      const texts = reviewer.snapshot().blocks.map(({ text }) => text);
      expect(texts.indexOf("After.")).toBe(texts.indexOf(item.text) + 1);
    });
});

describe("a block after a paragraph deleted whole that a join runs through", () => {
  /** "Alpha." merged into "Bravo.", "Bravo." deleted, then `insertion` after "Bravo.", tracked. */
  const joinThrough = async (
    session: HeadlessReviewerSession,
    insertion: (blockId: string) => FolioDocumentOperation,
  ) => {
    const reviewer = await openMarkdown(
      "Alpha.\n\nBravo.\n\nCharlie.",
      session,
    );
    const [alpha, bravo] = reviewer.snapshot().blocks;
    if (!alpha || !bravo) throw new Error("Missing fixture");
    for (const operation of [
      { id: "delete", type: "deleteBlock", blockId: bravo.id },
      {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: alpha.id,
        separator: " ",
      },
      insertion(bravo.id),
    ] satisfies FolioDocumentOperation[]) {
      expect(
        reviewer.applyOperations([operation], { mode: "tracked-changes" })
          .skipped,
      ).toEqual([]);
    }
    return reviewer;
  };
  const textsOf = (reviewer: FolioDocxReviewer) =>
    reviewer.snapshot().blocks.map(({ text }) => text);

  for (const session of SESSIONS)
    for (const placed of [
      "a paragraph after it",
      "a table after it",
    ] as const) {
      const insertion = insertions[placed];
      test(`takes ${placed} after the paragraph the join leaves (${session})`, async () => {
        const reviewer = await joinThrough(session, insertion);
        reviewer.acceptAll();
        const texts = textsOf(reviewer);
        const joined = texts.indexOf("Alpha. Charlie.");
        expect(joined).toBe(0);
        expect(texts.at(joined + 1)).toBe(
          placed === "a table after it" ? "Term" : "Inserted clause.",
        );
      });

      test(`rejects ${placed} to the document it started from (${session})`, async () => {
        const reviewer = await joinThrough(session, insertion);
        reviewer.rejectAll();
        expect(textsOf(reviewer)).toEqual(["Alpha.", "Bravo.", "Charlie."]);
      });
    }
});

describe("a suggested paragraph beside a pending join", () => {
  type Placed = {
    kind: string | undefined;
    sameList: boolean;
    order: readonly string[];
  };
  /**
   * A list item deleted, tracked, then its words' deletion rejected: the
   * break stays pending deletion, so the item joins the plain paragraph after
   * it. A paragraph put beside the item in `mode` lands as `Placed` says.
   */
  const placeBeside = async (
    session: HeadlessReviewerSession,
    type: "insertBeforeBlock" | "insertAfterBlock",
    mode: "suggested" | "tracked-changes",
  ): Promise<Placed> => {
    const reviewer = await open(session);
    const item = reviewer.snapshot().blocks.at(1);
    if (!item) throw new Error("Missing fixture");
    expect(
      reviewer.applyOperations(
        [{ id: "delete", type: "deleteBlock", blockId: item.id }],
        { mode: "tracked-changes" },
      ).skipped,
    ).toEqual([]);
    const words = reviewer
      .getChanges()
      .find(({ type: kind }) => kind === "deletion");
    if (!words) throw new Error("Missing deletion");
    expect(reviewer.rejectChange(words)).toBe(true);
    expect(
      reviewer.applyOperations(
        [{ id: "insert", type, blockId: item.id, text: "Inserted." }],
        { mode },
      ).skipped,
    ).toEqual([]);
    const blocks = reviewer.snapshot().blocks;
    const inserted = blocks.find(({ text }) => text === "Inserted.");
    return {
      kind: inserted?.kind,
      sameList: inserted?.listReference?.numId === item.listReference?.numId,
      order: blocks.map(({ text }) => text),
    };
  };

  for (const type of ["insertBeforeBlock", "insertAfterBlock"] as const) {
    test(`copies the paragraph it names as it stands (${type}, legacy)`, async () => {
      const placed = await placeBeside("legacy", type, "suggested");
      expect(placed.kind).toBe("listItem");
      expect(placed.sameList).toBe(true);
    });
    // Tracked and suggested insertions agree: accepting the suggestion leaves
    // what the tracked insertion's acceptance does.
    test(`lands where the tracked insertion does (${type}, canonical)`, async () => {
      expect(await placeBeside("canonical", type, "suggested")).toEqual(
        await placeBeside("canonical", type, "tracked-changes"),
      );
    });
  }
});

describe("a paragraph between a pending join and the table it reaches", () => {
  /**
   * "Bravo." deleted and "Alpha." then merged into it, tracked, before a
   * table: the join runs into a paragraph going whole. The legacy session
   * stops it at the table; on the canonical session accepting runs the
   * joined words into the table's first cell, so a paragraph put before
   * that cell goes before the whole join.
   */
  const beforeTable = async (
    order: "delete first" | "merge first" = "delete first",
    session?: HeadlessReviewerSession,
  ) => {
    const reviewer = await openMarkdown(
      "Alpha.\n\nBravo.\n\nCharlie.",
      session,
    );
    const [alpha, bravo] = reviewer.snapshot().blocks;
    if (!alpha || !bravo) throw new Error("Missing fixture");
    expect(
      reviewer.applyOperations(
        [{ id: "t", type: "insertTable", blockId: bravo.id, rows: [["Cell"]] }],
        { mode: "direct" },
      ).skipped,
    ).toEqual([]);
    const cell = reviewer.snapshot().blocks.find(({ text }) => text === "Cell");
    if (!cell) throw new Error("Missing table");
    const merge: FolioDocumentOperation = {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: alpha.id,
      separator: " ",
    };
    const deletion: FolioDocumentOperation = {
      id: "delete",
      type: "deleteBlock",
      blockId: bravo.id,
    };
    for (const operation of order === "delete first"
      ? [deletion, merge]
      : [merge, deletion]) {
      expect(
        reviewer.applyOperations([operation], { mode: "tracked-changes" })
          .skipped,
      ).toEqual([]);
    }
    return { reviewer, alpha: alpha.id, bravo: bravo.id, cell: cell.id };
  };

  const refusedOn = (anchor: "bravo" | "cell") =>
    anchor === "bravo" ? SESSIONS : (["legacy"] as const);
  for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
    for (const anchor of ["bravo", "cell"] as const) {
      for (const session of refusedOn(anchor))
        test(`is refused and changes nothing (${mode}, named by ${anchor}, ${session})`, async () => {
          const target = await beforeTable("delete first", session);
          const before = target.reviewer.snapshot();
          const result = target.reviewer.applyOperations(
            [
              {
                id: "insert",
                type:
                  anchor === "bravo" ? "insertAfterBlock" : "insertBeforeBlock",
                blockId: target[anchor],
                text: "Inserted.",
              },
            ],
            { mode },
          );
          expect(result.applied).toEqual([]);
          expect(result.skipped.map(({ reason }) => reason)).toEqual([
            "pendingParagraphMarkDeletion",
          ]);
          expect(target.reviewer.snapshot()).toEqual(before);
        });
    }
  }

  for (const mode of ["direct", "tracked-changes"] as const)
    test(`named by the cell, goes before the join, whose words run into the cell (${mode}, canonical)`, async () => {
      const { reviewer, cell } = await beforeTable("delete first", "canonical");
      expect(
        reviewer.applyOperations(
          [
            {
              id: "insert",
              type: "insertBeforeBlock",
              blockId: cell,
              text: "Inserted.",
            },
          ],
          { mode },
        ).skipped,
      ).toEqual([]);
      const rejected = await reviewerOf(await reviewer.toBuffer(), "canonical");
      rejected.rejectAll();
      expect(rejected.snapshot().blocks.map(({ text }) => text)).toEqual([
        ...(mode === "direct" ? ["Inserted."] : []),
        "Alpha.",
        "Bravo.",
        "Cell",
        "Charlie.",
      ]);
      reviewer.acceptAll();
      expect(
        reviewer.snapshot().blocks.map(({ text, table }) => ({
          text,
          inTable: table !== undefined,
        })),
      ).toEqual([
        { text: "Inserted.", inTable: false },
        { text: "Alpha. Cell", inTable: true },
        { text: "Charlie.", inTable: false },
      ]);
    });

  for (const mode of ["direct", "tracked-changes"] as const)
    test(`after the paragraph the join runs into, once it is deleted, is refused: accepting runs the joined words into the table (${mode}, canonical)`, async () => {
      // Deleting the merged paragraph deletes it whole, its break included
      // (survivor ids option C), so the joined words run into the table.
      const { reviewer, bravo } = await beforeTable("merge first", "canonical");
      const before = reviewer.snapshot();
      const result = reviewer.applyOperations(
        [
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: bravo,
            text: "Inserted.",
          },
        ],
        { mode },
      );
      expect(result.skipped.map(({ reason }) => reason)).toEqual([
        "pendingParagraphMarkDeletion",
      ]);
      expect(reviewer.snapshot()).toEqual(before);
      reviewer.acceptAll();
      expect(
        reviewer.snapshot().blocks.map(({ text, table }) => ({
          text,
          inTable: table !== undefined,
        })),
      ).toEqual([
        { text: "Alpha. Cell", inTable: true },
        { text: "Charlie.", inTable: false },
      ]);
    });

  test("is applied and stands apart once the paragraph the join runs into is deleted (legacy)", async () => {
    // Deleting the paragraph a join runs into keeps its break, the joined
    // paragraph's: nothing runs past it, so a paragraph after it stands.
    const { reviewer, bravo } = await beforeTable("merge first", "legacy");
    expect(
      reviewer.applyOperations(
        [
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: bravo,
            text: "Inserted.",
          },
        ],
        { mode: "tracked-changes" },
      ).skipped,
    ).toEqual([]);
    reviewer.acceptAll();
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual([
      "Alpha. ",
      "Inserted.",
      "Cell",
      "Charlie.",
    ]);
  });

  test("is applied and stands apart with a paragraph going whole before the table (legacy)", async () => {
    const { reviewer, alpha } = await beforeTable("delete first", "legacy");
    expect(
      reviewer.applyOperations(
        [
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: alpha,
            text: "Inserted.",
          },
        ],
        { mode: "tracked-changes" },
      ).skipped,
    ).toEqual([]);
    reviewer.acceptAll();
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual([
      "Alpha. ",
      "Inserted.",
      "Cell",
      "Charlie.",
    ]);
  });

  for (const mode of ["direct", "tracked-changes"] as const)
    test(`right after the words, is refused: accepting runs them into the table (${mode}, canonical)`, async () => {
      const { reviewer, alpha } = await beforeTable(
        "delete first",
        "canonical",
      );
      const before = reviewer.snapshot();
      const result = reviewer.applyOperations(
        [
          {
            id: "insert",
            type: "insertAfterBlock",
            blockId: alpha,
            text: "Inserted.",
          },
        ],
        { mode },
      );
      expect(result.skipped.map(({ reason }) => reason)).toEqual([
        "pendingParagraphMarkDeletion",
      ]);
      expect(reviewer.snapshot()).toEqual(before);
    });

  describe("when the table's first cell is pending", () => {
    const beforePendingTable = async (session: HeadlessReviewerSession) => {
      const reviewer = await openMarkdown(
        "Alpha.\n\nBravo.\n\nCharlie.",
        session,
      );
      const [alpha, bravo] = reviewer.snapshot().blocks;
      if (!alpha || !bravo) throw new Error("Missing fixture");
      for (const operation of [
        { id: "t", type: "insertTable", blockId: bravo.id, rows: [["Cell"]] },
        { id: "delete", type: "deleteBlock", blockId: bravo.id },
      ] satisfies FolioDocumentOperation[]) {
        expect(
          reviewer.applyOperations([operation], { mode: "tracked-changes" })
            .skipped,
        ).toEqual([]);
      }
      const merged = reviewer.applyOperations(
        [
          {
            id: "merge",
            type: "mergeBlockWithNext",
            blockId: alpha.id,
            separator: " ",
          },
        ],
        { mode: "tracked-changes" },
      );
      return { reviewer, alpha: alpha.id, bravo: bravo.id, merged };
    };

    // The canonical session refuses a join through a deleted paragraph into
    // a pending inserted table, so neither insertion has a join beside it.
    test("refuses the join into the pending table on the canonical session", async () => {
      const { reviewer, merged } = await beforePendingTable("canonical");
      expect(merged.applied).toEqual([]);
      expect(merged.skipped).toEqual([
        expect.objectContaining({
          id: "merge",
          message: expect.stringContaining("would run it into the table"),
        }),
      ]);
      expect(reviewer.snapshot().blocks.at(0)?.text).toBe("Alpha.");
    });

    test("right after the words, it is applied and stands apart", async () => {
      const { reviewer, alpha, merged } = await beforePendingTable("legacy");
      expect(merged.skipped).toEqual([]);
      expect(
        reviewer.applyOperations(
          [
            {
              id: "insert",
              type: "insertAfterBlock",
              blockId: alpha,
              text: "Inserted.",
            },
          ],
          { mode: "tracked-changes" },
        ).skipped,
      ).toEqual([]);
      reviewer.acceptAll();
      expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual([
        "Alpha. ",
        "Inserted.",
        "Cell",
        "Charlie.",
      ]);
    });

    test("past a paragraph going whole, it is refused", async () => {
      const { reviewer, bravo, merged } = await beforePendingTable("legacy");
      expect(merged.skipped).toEqual([]);
      expect(
        reviewer
          .applyOperations(
            [
              {
                id: "insert",
                type: "insertAfterBlock",
                blockId: bravo,
                text: "Inserted.",
              },
            ],
            { mode: "tracked-changes" },
          )
          .skipped.map(({ reason }) => reason),
      ).toEqual(["pendingParagraphMarkDeletion"]);
    });
  });

  for (const session of SESSIONS)
    test(`is applied once the join is rejected (${session})`, async () => {
      const { reviewer, cell } = await beforeTable("delete first", session);
      const join = reviewer
        .getChanges()
        .find(({ type }) => type === "paragraphMarkDeleted");
      if (!join) throw new Error("Missing join");
      expect(reviewer.rejectChange(join)).toBe(true);
      expect(
        reviewer.applyOperations(
          [
            {
              id: "insert",
              type: "insertBeforeBlock",
              blockId: cell,
              text: "Inserted.",
            },
          ],
          { mode: "tracked-changes" },
        ).skipped,
      ).toEqual([]);
      reviewer.acceptAll();
      expect(reviewer.snapshot().blocks.map(({ text }) => text)).toContain(
        "Inserted.",
      );
    });
});

const EMPTY_SOURCE_MARKDOWN = "Alpha.\n\nBravo.\n\nCharlie.";

type FromEmptySourceOptions = {
  session: HeadlessReviewerSession;
  mode: "direct" | "tracked-changes";
  batches: (ids: {
    empty: string;
    bravo: string;
  }) => FolioDocumentOperation[][];
};

/** An empty paragraph after "Alpha.", merged into "Bravo.", then `batches` in `mode`. */
const fromEmptySource = async ({
  session,
  mode,
  batches,
}: FromEmptySourceOptions) => {
  const reviewer = await openMarkdown(EMPTY_SOURCE_MARKDOWN, session);
  const alpha = reviewer.snapshot().blocks.at(0);
  if (!alpha) throw new Error("Missing fixture");
  reviewer.applyOperations(
    [{ id: "empty", type: "insertAfterBlock", blockId: alpha.id, text: "" }],
    { mode: "direct" },
  );
  const [, empty, bravo] = reviewer.snapshot().blocks;
  if (!empty || !bravo) throw new Error("Missing fixture");
  const merged = reviewer.applyOperations(
    [{ id: "merge", type: "mergeBlockWithNext", blockId: empty.id }],
    { mode },
  );
  const mergedParagraphs = merged.applied.find(
    ({ paragraphIdentity }) => paragraphIdentity !== undefined,
  )?.paragraphIdentity;
  for (const batch of batches({
    empty:
      mergedParagraphs?.type === "joined" &&
      mergedParagraphs.retiredBlockId === empty.id
        ? mergedParagraphs.blockId
        : empty.id,
    bravo: bravo.id,
  })) {
    expect(reviewer.applyOperations(batch, { mode }).skipped).toEqual([]);
  }
  if (mode !== "direct") reviewer.acceptAll();
  return reviewer.snapshot().blocks.map(({ text }) => text);
};

describe("a pending join from an empty paragraph", () => {
  for (const session of SESSIONS) {
    for (const [placed, insertion] of Object.entries(insertions)) {
      if (placed.endsWith("before it")) continue;
      test(`takes ${placed} after the paragraph it leaves, as direct edits do (${session})`, async () => {
        const batches = ({ empty }: { empty: string }) => [[insertion(empty)]];
        const direct = await fromEmptySource({
          session,
          mode: "direct",
          batches,
        });
        const bravo = direct.indexOf("Bravo.");
        const inserted = direct.findIndex(
          (text) => text === "Inserted clause." || text === "Term",
        );
        expect(bravo).toBeGreaterThan(-1);
        expect(inserted).toBeGreaterThan(-1);
        expect(bravo).toBeLessThan(inserted);
        expect(
          await fromEmptySource({ session, mode: "tracked-changes", batches }),
        ).toEqual(direct);
      });
    }
    for (const order of ["after first", "before first"] as const) {
      test(`keeps insertions on both sides in a batch as one at a time (${order}, ${session})`, async () => {
        const operations = ({
          empty,
          bravo,
        }: {
          empty: string;
          bravo: string;
        }) => {
          const after: FolioDocumentOperation = {
            id: "after",
            type: "insertAfterBlock",
            blockId: empty,
            text: "After.",
          };
          const before: FolioDocumentOperation = {
            id: "before",
            type: "insertBeforeBlock",
            blockId: bravo,
            text: "Before.",
          };
          return order === "after first" ? [after, before] : [before, after];
        };
        const sequential = await fromEmptySource({
          session,
          mode: "tracked-changes",
          batches: (ids) => operations(ids).map((operation) => [operation]),
        });
        expect(
          await fromEmptySource({
            session,
            mode: "tracked-changes",
            batches: (ids) => [operations(ids)],
          }),
        ).toEqual(sequential);
        expect(sequential).toEqual([
          "Alpha.",
          "Before.",
          "Bravo.",
          "After.",
          "Charlie.",
        ]);
      });
    }
  }
});

const SPLIT_MARKDOWN =
  "Alpha.\n\nThe Buyer pays each invoice within thirty days.\n\nSigned in two copies.\n\nCharlie.";
const SPLIT_SURVIVOR = "Signed in two copies.";

/**
 * How the clause comes to join the paragraph after it: a tracked merge
 * leaves its words, and a tracked deletion of which only the words are
 * accepted leaves it empty with its break pending deletion.
 */
const joinSources = {
  "a tracked merge": "merge",
  "an emptied deletion": "emptied",
} as const;

type JoinSource = (typeof joinSources)[keyof typeof joinSources];

type SplitFixture = {
  reviewer: FolioDocxReviewer;
  clause: string;
  survivor: string;
};

/** The clause joining "Signed in two copies." as `source` makes it. */
const openSplitFixture = async (
  source: JoinSource,
  session: HeadlessReviewerSession,
): Promise<SplitFixture> => {
  const reviewer = await openMarkdown(SPLIT_MARKDOWN, session);
  const [, clause, survivor] = reviewer.snapshot().blocks;
  if (!clause || survivor?.text !== SPLIT_SURVIVOR) {
    throw new Error("Missing fixture");
  }
  const operation: FolioDocumentOperation =
    source === "merge"
      ? {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: clause.id,
          separator: " ",
        }
      : { id: "delete", type: "deleteBlock", blockId: clause.id };
  expect(
    reviewer.applyOperations([operation], { mode: "tracked-changes" }).skipped,
  ).toEqual([]);
  if (source === "emptied") {
    const words = reviewer
      .getChanges()
      .find(
        (change) => change.type === "deletion" && change.text === clause.text,
      );
    if (!words) throw new Error("Missing the clause's deleted words");
    expect(reviewer.acceptChange(words)).toBe(true);
    // Its break stays pending deletion: the clause reads as part of the next paragraph.
    expect(reviewer.getChanges().length).toBeGreaterThan(0);
  }
  return { reviewer, clause: clause.id, survivor: survivor.id };
};

const splitInsertions = {
  "a paragraph": (blockId: string): FolioDocumentOperation => ({
    id: "insert",
    type: "insertAfterBlock",
    blockId,
    text: "Inserted clause.",
  }),
  "a table": insertions["a table after it"],
};

type SplitBatchOptions = {
  session: HeadlessReviewerSession;
  source: JoinSource;
  insertion: (blockId: string) => FolioDocumentOperation;
  /** Whether the block is put after the joining clause or after the paragraph it joins. */
  anchor: "clause" | "survivor";
  offset: number;
  order: "insertion first" | "split first";
  resolution: "accept" | "reject";
};

/** One batch: the block put after `anchor`, and the joined paragraph split at `offset`. */
const splitBatch = async ({
  session,
  source,
  insertion,
  anchor,
  offset,
  order,
  resolution,
}: SplitBatchOptions): Promise<string[]> => {
  const fixture = await openSplitFixture(source, session);
  const insert = insertion(fixture[anchor]);
  const split: FolioDocumentOperation = {
    id: "split",
    type: "splitBlock",
    blockId: fixture.survivor,
    offset,
  };
  expect(
    fixture.reviewer.applyOperations(
      order === "insertion first" ? [insert, split] : [split, insert],
      { mode: "tracked-changes" },
    ).skipped,
  ).toEqual([]);
  if (resolution === "accept") fixture.reviewer.acceptAll();
  else fixture.reviewer.rejectAll();
  return fixture.reviewer.snapshot().blocks.map(({ text }) => text);
};

// A batch reads the document as it stands before the batch, so a block put
// after a paragraph that reads as part of the next one lands after the
// paragraph it joins as the batch read it, even when the same batch splits
// that paragraph, whichever it names first.
describe("a block after a paragraph that joins one the batch splits", () => {
  for (const session of SESSIONS)
    for (const [sourceName, source] of Object.entries(joinSources)) {
      for (const [inserted, insertion] of Object.entries(splitInsertions)) {
        for (const offset of [1, 14, SPLIT_SURVIVOR.length - 1]) {
          for (const order of ["insertion first", "split first"] as const) {
            const options = { session, source, insertion, offset, order };
            test(`follows the whole joined paragraph, as one put after it does (${sourceName}, ${inserted}, split at ${offset}, ${order}, ${session})`, async () => {
              const accepted = await splitBatch({
                ...options,
                anchor: "clause",
                resolution: "accept",
              });
              // On the canonical session a paragraph after the clause, which
              // keeps its words, goes right after it and takes them; the
              // paragraph it joined stands, split as the batch splits it.
              if (
                session === "canonical" &&
                source === "merge" &&
                inserted === "a paragraph"
              ) {
                expect(accepted).toEqual([
                  "Alpha.",
                  "The Buyer pays each invoice within thirty days. Inserted clause.",
                  SPLIT_SURVIVOR.slice(0, offset),
                  SPLIT_SURVIVOR.slice(offset),
                  "Charlie.",
                ]);
                return;
              }
              const first = accepted.findIndex(
                (text) => text === "Inserted clause." || text === "Term",
              );
              expect(accepted.at(first - 1)).toBe(SPLIT_SURVIVOR.slice(offset));
              expect(accepted).toEqual(
                await splitBatch({
                  ...options,
                  anchor: "survivor",
                  resolution: "accept",
                }),
              );
            });
            test(`rejects to the document it started from (${sourceName}, ${inserted}, split at ${offset}, ${order}, ${session})`, async () => {
              const fixture = await openSplitFixture(source, session);
              fixture.reviewer.rejectAll();
              expect(
                await splitBatch({
                  ...options,
                  anchor: "clause",
                  resolution: "reject",
                }),
              ).toEqual(
                fixture.reviewer.snapshot().blocks.map(({ text }) => text),
              );
            });
          }
        }
      }
    }
});

describe("a block before a pending join whose first paragraph the batch deletes", () => {
  for (const session of SESSIONS) {
    const read = async (batched: boolean, resolution: "accept" | "reject") => {
      const reviewer = await open(session);
      const [heading, item] = reviewer.snapshot().blocks;
      if (!heading || !item) throw new Error("Missing fixture");
      reviewer.applyOperations(
        [
          {
            id: "merge",
            type: "mergeBlockWithNext",
            blockId: heading.id,
            separator: " ",
          },
        ],
        { mode: "tracked-changes" },
      );
      // One at a time, the insertion runs first: it is stated further into the document.
      const operations: FolioDocumentOperation[] = [
        {
          id: "before",
          type: "insertBeforeBlock",
          blockId: item.id,
          text: "Before.",
        },
        { id: "delete", type: "deleteBlock", blockId: heading.id },
      ];
      for (const batch of batched
        ? [operations]
        : operations.map((operation) => [operation])) {
        expect(
          reviewer.applyOperations(batch, { mode: "tracked-changes" }).skipped,
        ).toEqual([]);
      }
      if (resolution === "accept") reviewer.acceptAll();
      else reviewer.rejectAll();
      return reviewer.snapshot().blocks.map(({ text }) => text);
    };

    test(`rejects to the document it started from (${session})`, async () => {
      const original = (await open(session))
        .snapshot()
        .blocks.map(({ text }) => text);
      expect(await read(true, "reject")).toEqual(original);
    });

    test(`accepts in a batch as one at a time (${session})`, async () => {
      expect(await read(true, "accept")).toEqual(await read(false, "accept"));
    });
  }
});

describe("a block after a pending join whose paragraph the batch splits, inside a comment", () => {
  // See "a block after a paragraph that joins one the batch splits".
  for (const session of SESSIONS)
    test(`carries the comments a block put after that paragraph carries (${session})`, async () => {
      const read = async (anchor: "heading" | "item") => {
        const reviewer = await open(session);
        const [heading, item] = reviewer.snapshot().blocks;
        if (!heading || !item) throw new Error("Missing fixture");
        for (const operation of [
          {
            id: "comment",
            type: "commentOnBlock",
            blockId: item.id,
            comment: { text: "Review this clause." },
          },
          {
            id: "merge",
            type: "mergeBlockWithNext",
            blockId: heading.id,
            separator: " ",
          },
        ] satisfies FolioDocumentOperation[]) {
          expect(
            reviewer.applyOperations([operation], { mode: "tracked-changes" })
              .skipped,
          ).toEqual([]);
        }
        expect(
          reviewer.applyOperations(
            [
              { id: "split", type: "splitBlock", blockId: item.id, offset: 8 },
              insertions["a table after it"](
                anchor === "heading" ? heading.id : item.id,
              ),
            ],
            { mode: "tracked-changes" },
          ).skipped,
        ).toEqual([]);
        return reviewer.getComments().map(({ anchoredText }) => anchoredText);
      };
      const afterItem = await read("item");
      expect(afterItem.length).toBeGreaterThan(0);
      expect(await read("heading")).toEqual(afterItem);
    });
});
describe("a table and a paragraph inserted before a deleted final paragraph", () => {
  const beforeFinal = async (
    mode: "direct" | "tracked-changes",
    session?: HeadlessReviewerSession,
  ) => {
    const reviewer = await openMarkdown(
      "Alpha.\n\nBravo.\n\nCharlie.",
      session,
    );
    const [, bravo, charlie] = reviewer.snapshot().blocks;
    if (!bravo || !charlie) throw new Error("Missing fixture");
    const joinedIds = new Map<string, string>();
    for (const operations of [
      [{ id: "delete", type: "deleteBlock", blockId: charlie.id }],
      [
        {
          id: "table",
          type: "insertTable",
          blockId: bravo.id,
          rows: [["Cell"]],
        },
        {
          id: "after",
          type: "insertAfterBlock",
          blockId: bravo.id,
          text: "After.",
        },
      ],
    ] satisfies FolioDocumentOperation[][]) {
      const result = reviewer.applyOperations(
        operations.map((operation) =>
          Object.assign({}, operation, {
            blockId: joinedIds.get(operation.blockId) ?? operation.blockId,
          }),
        ),
        { mode },
      );
      // A canonical tracked table after a paragraph whose break a pending
      // final-paragraph deletion retired is refused, as a join into a pending
      // inserted table is.
      const refused =
        session === "canonical" &&
        mode === "tracked-changes" &&
        operations.length > 1;
      expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual(
        refused
          ? [{ id: "table", reason: "pendingParagraphMarkDeletion" }]
          : [],
      );
      for (const { paragraphIdentity: joined } of result.applied) {
        if (joined?.type === "joined")
          joinedIds.set(joined.retiredBlockId, joined.blockId);
      }
    }
    return reviewer;
  };
  const textsOf = (reviewer: FolioDocxReviewer) =>
    reviewer.snapshot().blocks.map(({ text }) => text);

  test("leaves no empty paragraph once accepted, as direct edits leave none", async () => {
    const direct = await beforeFinal("direct");
    const tracked = await beforeFinal("tracked-changes");
    tracked.acceptAll();
    expect(textsOf(direct)).toEqual(["Alpha.", "Bravo.", "Cell", "After."]);
    expect(textsOf(tracked)).toEqual(textsOf(direct));
  });

  for (const session of SESSIONS)
    test(`rejects to the document it started from, before and after a reopen (${session})`, async () => {
      const tracked = await beforeFinal("tracked-changes", session);
      const reopened = await reviewerOf(await tracked.toBuffer(), session);
      for (const reviewer of [tracked, reopened]) {
        reviewer.rejectAll();
        expect(textsOf(reviewer)).toEqual(["Alpha.", "Bravo.", "Charlie."]);
      }
    });
});
