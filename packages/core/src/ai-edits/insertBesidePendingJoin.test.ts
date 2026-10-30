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
