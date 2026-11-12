/**
 * A block inserted beside a paragraph whose break a tracked merge removes
 * treats the two as the joined paragraph they become: a paragraph copies the
 * properties the joined one will have, and a block after it goes after the
 * paragraph the join leaves. Accepting gives what the same edits applied
 * directly give.
 */

import { describe, expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const MARKDOWN = "# Terms\n\n1. Numbered clause one\n\nUnnumbered body text.";

const open = async () => {
  const bytes = (await ensureParaIds(await createDocx(fromMarkdown(MARKDOWN))))
    .docx;
  return FolioDocxReviewer.fromBuffer(bytes, { author: "Test" });
};

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
  mode: "direct" | "tracked-changes";
  batched: boolean;
  merged: (numId: number) => MergedProperties;
  insertion: (blockId: string) => FolioDocumentOperation;
};

/** Merge the heading into the list item, insert beside it, accept, and read. */
const edit = async ({ mode, batched, merged, insertion }: EditOptions) => {
  const reviewer = await open();
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
  const insert = insertion(heading.id);
  const batches = batched ? [[merge, insert]] : [[merge], [insert]];
  for (const operations of batches) {
    const { skipped } = reviewer.applyOperations(operations, {
      mode,
      atomic: true,
    });
    expect(skipped).toEqual([]);
  }
  if (mode !== "direct") reviewer.acceptAll();
  const reopened = await FolioDocxReviewer.fromBuffer(
    await reviewer.toBuffer(),
    {
      author: "Test",
    },
  );
  return reopened.snapshot().blocks.map((block) => ({
    text: block.text,
    styleId: block.styleId,
    statedNumbering: block.statedNumbering,
    listReference: block.listReference,
    table: block.table,
  }));
};

describe("an insertion beside a paragraph a tracked merge joins", () => {
  for (const [placed, insertion] of Object.entries(insertions)) {
    for (const [name, merged] of Object.entries(requests)) {
      // One batch cannot both join a paragraph and place a block at its end.
      for (const batched of placed.endsWith("before it")
        ? [true, false]
        : [false]) {
        test(`reads as the joined paragraph (${placed}, ${name}, ${batched ? "one batch" : "two batches"})`, async () => {
          const direct = await edit({
            mode: "direct",
            batched,
            merged,
            insertion,
          });
          expect(
            direct.some(
              ({ text }) => text.startsWith("Inserted") || text === "Term",
            ),
          ).toBe(true);
          expect(
            await edit({ mode: "tracked-changes", batched, merged, insertion }),
          ).toEqual(direct);
        });
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
  test("after a join through a paragraph deleted whole goes where the direct edits put it", async () => {
    const read = async (mode: "direct" | "tracked-changes") => {
      const reviewer = await open();
      const [heading, item] = reviewer.snapshot().blocks;
      if (!heading || !item) throw new Error("Missing fixture");
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
        expect(reviewer.applyOperations([operation], { mode }).skipped).toEqual(
          [],
        );
      }
      if (mode !== "direct") reviewer.acceptAll();
      return reviewer.snapshot().blocks.map(({ text }) => text);
    };
    const direct = await read("direct");
    expect(direct).toEqual(["Terms Unnumbered body text.", "After."]);
    expect(await read("tracked-changes")).toEqual(direct);
  });
});

describe("an insertion after a paragraph whose break a final deletion retires", () => {
  test("stays beside it: the emptied final paragraph has no words to join", async () => {
    const reviewer = await open();
    const blocks = reviewer.snapshot().blocks;
    const [item, last] = [blocks.at(1), blocks.at(-1)];
    if (!item || !last) throw new Error("Missing fixture");
    reviewer.applyOperations(
      [{ id: "delete", type: "deleteBlock", blockId: last.id }],
      {
        mode: "tracked-changes",
      },
    );
    reviewer.applyOperations(
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
    const texts = reviewer.snapshot().blocks.map(({ text }) => text);
    expect(texts.indexOf("After.")).toBe(texts.indexOf(item.text) + 1);
  });
});

describe("a block after a paragraph deleted whole that a join runs through", () => {
  /** "Alpha." merged into "Bravo.", "Bravo." deleted, then `insertion` after "Bravo.", tracked. */
  const joinThrough = async (
    insertion: (blockId: string) => FolioDocumentOperation,
  ) => {
    const bytes = (
      await ensureParaIds(
        await createDocx(fromMarkdown("Alpha.\n\nBravo.\n\nCharlie.")),
      )
    ).docx;
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
      author: "Test",
    });
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

  for (const placed of ["a paragraph after it", "a table after it"] as const) {
    const insertion = insertions[placed];
    test(`takes ${placed} after the paragraph the join leaves`, async () => {
      const reviewer = await joinThrough(insertion);
      reviewer.acceptAll();
      const texts = textsOf(reviewer);
      const joined = texts.indexOf("Alpha. Charlie.");
      expect(joined).toBe(0);
      expect(texts.at(joined + 1)).toBe(
        placed === "a table after it" ? "Term" : "Inserted clause.",
      );
    });

    test(`rejects ${placed} to the document it started from`, async () => {
      const reviewer = await joinThrough(insertion);
      reviewer.rejectAll();
      expect(textsOf(reviewer)).toEqual(["Alpha.", "Bravo.", "Charlie."]);
    });
  }
});

describe("a suggested paragraph beside a pending join", () => {
  for (const type of ["insertBeforeBlock", "insertAfterBlock"] as const) {
    test(`copies the paragraph it names as it stands (${type})`, async () => {
      const reviewer = await open();
      const item = reviewer.snapshot().blocks.at(1);
      if (!item) throw new Error("Missing fixture");
      expect(
        reviewer.applyOperations(
          [{ id: "delete", type: "deleteBlock", blockId: item.id }],
          { mode: "tracked-changes" },
        ).skipped,
      ).toEqual([]);
      // Rejecting the words alone leaves the break pending deletion: the list
      // item now joins the plain paragraph after it.
      const words = reviewer
        .getChanges()
        .find(({ type: kind }) => kind === "deletion");
      if (!words) throw new Error("Missing deletion");
      expect(reviewer.rejectChange(words)).toBe(true);
      expect(
        reviewer.applyOperations(
          [{ id: "insert", type, blockId: item.id, text: "Inserted." }],
          { mode: "suggested" },
        ).skipped,
      ).toEqual([]);
      const inserted = reviewer
        .snapshot()
        .blocks.find(({ text }) => text === "Inserted.");
      expect(inserted?.kind).toBe("listItem");
      expect(inserted?.listReference?.numId).toBe(item.listReference?.numId);
    });
  }
});

describe("a paragraph between a pending join and the table that stops it", () => {
  /**
   * "Bravo." deleted and "Alpha." then merged into it, tracked, before a
   * table: the join runs into a paragraph going whole, and the table stops it.
   */
  const beforeTable = async (
    order: "delete first" | "merge first" = "delete first",
  ) => {
    const bytes = (
      await ensureParaIds(
        await createDocx(fromMarkdown("Alpha.\n\nBravo.\n\nCharlie.")),
      )
    ).docx;
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
      author: "Test",
    });
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

  for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
    for (const anchor of ["bravo", "cell"] as const) {
      test(`is refused and changes nothing (${mode}, named by ${anchor})`, async () => {
        const target = await beforeTable();
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

  test("is applied and stands apart once the paragraph the join runs into is deleted", async () => {
    // Deleting the paragraph a join runs into keeps its break, the joined
    // paragraph's: nothing runs past it, so a paragraph after it stands.
    const { reviewer, bravo } = await beforeTable("merge first");
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

  test("is applied and stands apart with a paragraph going whole before the table", async () => {
    const { reviewer, alpha } = await beforeTable();
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

  describe("when the table's first cell is pending", () => {
    const beforePendingTable = async () => {
      const bytes = (
        await ensureParaIds(
          await createDocx(fromMarkdown("Alpha.\n\nBravo.\n\nCharlie.")),
        )
      ).docx;
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
        author: "Test",
      });
      const [alpha, bravo] = reviewer.snapshot().blocks;
      if (!alpha || !bravo) throw new Error("Missing fixture");
      for (const operation of [
        { id: "t", type: "insertTable", blockId: bravo.id, rows: [["Cell"]] },
        { id: "delete", type: "deleteBlock", blockId: bravo.id },
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: alpha.id,
          separator: " ",
        },
      ] satisfies FolioDocumentOperation[]) {
        expect(
          reviewer.applyOperations([operation], { mode: "tracked-changes" })
            .skipped,
        ).toEqual([]);
      }
      return { reviewer, alpha: alpha.id, bravo: bravo.id };
    };

    test("right after the words, it is applied and stands apart", async () => {
      const { reviewer, alpha } = await beforePendingTable();
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
      const { reviewer, bravo } = await beforePendingTable();
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

  test("is applied once the join is rejected", async () => {
    const { reviewer, cell } = await beforeTable();
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

/** An empty paragraph after "Alpha.", merged into "Bravo.", then `edit` in `mode`. */
const fromEmptySource = async (
  mode: "direct" | "tracked-changes",
  batches: (ids: {
    empty: string;
    bravo: string;
  }) => FolioDocumentOperation[][],
) => {
  const bytes = (
    await ensureParaIds(await createDocx(fromMarkdown(EMPTY_SOURCE_MARKDOWN)))
  ).docx;
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
    author: "Test",
  });
  const alpha = reviewer.snapshot().blocks.at(0);
  if (!alpha) throw new Error("Missing fixture");
  reviewer.applyOperations(
    [{ id: "empty", type: "insertAfterBlock", blockId: alpha.id, text: "" }],
    { mode: "direct" },
  );
  const [, empty, bravo] = reviewer.snapshot().blocks;
  if (!empty || !bravo) throw new Error("Missing fixture");
  reviewer.applyOperations(
    [{ id: "merge", type: "mergeBlockWithNext", blockId: empty.id }],
    { mode },
  );
  for (const batch of batches({ empty: empty.id, bravo: bravo.id })) {
    expect(reviewer.applyOperations(batch, { mode }).skipped).toEqual([]);
  }
  if (mode !== "direct") reviewer.acceptAll();
  return reviewer.snapshot().blocks.map(({ text }) => text);
};

describe("a pending join from an empty paragraph", () => {
  for (const [placed, insertion] of Object.entries(insertions)) {
    if (placed.endsWith("before it")) continue;
    test(`takes ${placed} after the paragraph it leaves, as direct edits do`, async () => {
      const batches = ({ empty }: { empty: string }) => [[insertion(empty)]];
      const direct = await fromEmptySource("direct", batches);
      const bravo = direct.indexOf("Bravo.");
      const inserted = direct.findIndex(
        (text) => text === "Inserted clause." || text === "Term",
      );
      expect(bravo).toBeGreaterThan(-1);
      expect(inserted).toBeGreaterThan(-1);
      expect(bravo).toBeLessThan(inserted);
      expect(await fromEmptySource("tracked-changes", batches)).toEqual(direct);
    });
  }
  for (const order of ["after first", "before first"] as const) {
    test(`keeps insertions on both sides in a batch as one at a time (${order})`, async () => {
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
      const sequential = await fromEmptySource("tracked-changes", (ids) =>
        operations(ids).map((operation) => [operation]),
      );
      expect(
        await fromEmptySource("tracked-changes", (ids) => [operations(ids)]),
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
});

describe("a block after a pending join whose paragraph the batch splits", () => {
  test("follows the half the join runs into, in a batch as one at a time", async () => {
    const read = async (batched: boolean) => {
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
      // One at a time, the split runs first too: it edits further into the document.
      const operations: FolioDocumentOperation[] = [
        { id: "split", type: "splitBlock", blockId: item.id, offset: 8 },
        insertions["a table after it"](heading.id),
      ];
      for (const batch of batched
        ? [operations]
        : operations.map((operation) => [operation])) {
        expect(
          reviewer.applyOperations(batch, { mode: "tracked-changes" }).skipped,
        ).toEqual([]);
      }
      reviewer.acceptAll();
      return reviewer.snapshot().blocks.map(({ text }) => text);
    };
    const sequential = await read(false);
    expect(sequential.slice(0, 2)).toEqual(["Terms Numbered", "Term"]);
    expect(await read(true)).toEqual(sequential);
  });
});

describe("a block before a pending join whose first paragraph the batch deletes", () => {
  const read = async (batched: boolean, resolution: "accept" | "reject") => {
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

  test("rejects to the document it started from", async () => {
    const original = (await open()).snapshot().blocks.map(({ text }) => text);
    expect(await read(true, "reject")).toEqual(original);
  });

  test("accepts in a batch as one at a time", async () => {
    expect(await read(true, "accept")).toEqual(await read(false, "accept"));
  });
});

describe("a block after a pending join whose paragraph the batch splits, inside a comment", () => {
  test("carries the comment running across the boundary it lands on, in a batch as one at a time", async () => {
    const read = async (batched: boolean) => {
      const reviewer = await open();
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
      // One at a time, the split runs first: it edits further into the document.
      const operations: FolioDocumentOperation[] = [
        { id: "split", type: "splitBlock", blockId: item.id, offset: 8 },
        insertions["a table after it"](heading.id),
      ];
      for (const batch of batched
        ? [operations]
        : operations.map((operation) => [operation])) {
        expect(
          reviewer.applyOperations(batch, { mode: "tracked-changes" }).skipped,
        ).toEqual([]);
      }
      return reviewer.getComments().map(({ anchoredText }) => anchoredText);
    };
    const sequential = await read(false);
    expect(sequential.some((text) => text.includes("Term"))).toBe(true);
    expect(await read(true)).toEqual(sequential);
  });
});

describe("a table and a paragraph inserted before a deleted final paragraph", () => {
  const beforeFinal = async (mode: "direct" | "tracked-changes") => {
    const bytes = (
      await ensureParaIds(
        await createDocx(fromMarkdown("Alpha.\n\nBravo.\n\nCharlie.")),
      )
    ).docx;
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
      author: "Test",
    });
    const [, bravo, charlie] = reviewer.snapshot().blocks;
    if (!bravo || !charlie) throw new Error("Missing fixture");
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
      expect(reviewer.applyOperations(operations, { mode }).skipped).toEqual(
        [],
      );
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

  test("rejects to the document it started from, before and after a reopen", async () => {
    const tracked = await beforeFinal("tracked-changes");
    const reopened = await FolioDocxReviewer.fromBuffer(
      await tracked.toBuffer(),
      { author: "Test" },
    );
    for (const reviewer of [tracked, reopened]) {
      reviewer.rejectAll();
      expect(textsOf(reviewer)).toEqual(["Alpha.", "Bravo.", "Charlie."]);
    }
  });
});
