/**
 * List labels follow the document as it stands. Every reader of a reviewer
 * (`getContent()`, `getContentAsText()`, Markdown of `toDocument()`) reads
 * the number an item shows after an operation adds, removes, re-levels or
 * restarts list items, and a save and reopen reads the same numbers.
 */

import { describe, expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { toMarkdown } from "../markdown";
import { FolioDocxReviewer } from "./headless";
import { isFolioAIContentBlock } from "./snapshot";

const LIST = [
  "Payment happens in stages.",
  "1. Deposit on signature\n2. Balance on delivery\n3. Retention after inspection",
  "Closing remarks.",
].join("\n\n");

const reviewerOf = async (markdown: string): Promise<FolioDocxReviewer> => {
  const { docx } = await ensureParaIds(await createDocx(fromMarkdown(markdown)));
  return FolioDocxReviewer.fromBuffer(docx, { author: "Agent" });
};

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) {
    throw new Error(`no block reads "${text}"`);
  }
  return block.id;
};

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
  operations: FolioDocumentOperation[],
): void => {
  const result = reviewer.applyDocumentOperations({ version: 1, mode, operations });
  expect(result.issues).toEqual([]);
  expect(result.status).toBe("committed");
};

/** `getContent()`: `1. Deposit on signature`, or the text alone. */
const contentLabels = (reviewer: FolioDocxReviewer): string[] =>
  reviewer
    .getContent()
    .filter(isFolioAIContentBlock)
    .map(({ displayLabel, text }) => `${displayLabel ?? ""} ${text}`.trim());

/** `getContentAsText()` without the block ids. */
const textLabels = (reviewer: FolioDocxReviewer): string[] =>
  reviewer
    .getContentAsText()
    .split("\n")
    .map((line) => line.replace(/^\[[^\]]+\]\s*/u, ""));

/** Markdown of the reviewer's current document, one line per block. */
const markdownLabels = (reviewer: FolioDocxReviewer): string[] =>
  toMarkdown(reviewer.toDocument(), {
    annotations: "strip",
    trackedChanges: "clean",
    comments: "strip",
  })
    .split("\n")
    .map((line) => line.replaceAll("\\", "").trim())
    .filter((line) => line.length > 0);

/** Every reader of the reviewer, and of its saved package, reads `expected`. */
const expectLabels = async (reviewer: FolioDocxReviewer, expected: string[]): Promise<void> => {
  expect(contentLabels(reviewer)).toEqual(expected);
  expect(textLabels(reviewer)).toEqual(expected);
  expect(markdownLabels(reviewer)).toEqual(expected);
  const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  expect(contentLabels(reopened)).toEqual(expected);
};

describe("list labels after an operation", () => {
  test("an inserted item takes its number and the items after it renumber", async () => {
    const reviewer = await reviewerOf(LIST);
    apply(reviewer, "direct", [
      {
        id: "1",
        type: "insertAfterBlock",
        blockId: blockId(reviewer, "Deposit on signature"),
        text: "Interim payment",
      },
    ]);

    await expectLabels(reviewer, [
      "Payment happens in stages.",
      "1. Deposit on signature",
      "2. Interim payment",
      "3. Balance on delivery",
      "4. Retention after inspection",
      "Closing remarks.",
    ]);
  });

  test("the items after a deleted item renumber", async () => {
    const reviewer = await reviewerOf(LIST);
    apply(reviewer, "direct", [
      { id: "1", type: "deleteBlock", blockId: blockId(reviewer, "Deposit on signature") },
    ]);

    await expectLabels(reviewer, [
      "Payment happens in stages.",
      "1. Balance on delivery",
      "2. Retention after inspection",
      "Closing remarks.",
    ]);
  });

  test("an item moved a level down counts at that level, and its siblings close up", async () => {
    const reviewer = await reviewerOf(
      [
        "Payment happens in stages.",
        "1. Deposit on signature\n   1. Within a week\n2. Balance on delivery\n3. Retention after inspection",
        "Closing remarks.",
      ].join("\n\n"),
    );
    const numId = reviewer.getContent().find(({ text }) => text === "Balance on delivery")
      ?.listReference?.numId;
    expect(numId).toBeDefined();
    apply(reviewer, "direct", [
      {
        id: "1",
        type: "setBlockParagraphProperties",
        blockId: blockId(reviewer, "Balance on delivery"),
        properties: { numbering: { numId: numId ?? 0, level: 1 } },
      },
    ]);

    await expectLabels(reviewer, [
      "Payment happens in stages.",
      "1. Deposit on signature",
      "1. Within a week",
      "2. Balance on delivery",
      "2. Retention after inspection",
      "Closing remarks.",
    ]);
  });

  test("an item that starts a new list restarts the count", async () => {
    const reviewer = await reviewerOf(LIST);
    apply(reviewer, "direct", [
      {
        id: "1",
        type: "setBlockParagraphProperties",
        blockId: blockId(reviewer, "Balance on delivery"),
        properties: { numbering: { start: "new", kind: "numbered" } },
      },
    ]);

    await expectLabels(reviewer, [
      "Payment happens in stages.",
      "1. Deposit on signature",
      "1. Balance on delivery",
      // The rest of the first list continues it.
      "2. Retention after inspection",
      "Closing remarks.",
    ]);
  });

  test("a tracked insertion is numbered in the reviewer as in the saved package", async () => {
    const reviewer = await reviewerOf(LIST);
    apply(reviewer, "tracked-changes", [
      {
        id: "1",
        type: "insertBeforeBlock",
        blockId: blockId(reviewer, "Deposit on signature"),
        text: "Reservation fee",
      },
    ]);

    const expected = [
      "Payment happens in stages.",
      "1. Reservation fee",
      "2. Deposit on signature",
      "3. Balance on delivery",
      "4. Retention after inspection",
      "Closing remarks.",
    ];
    expect(contentLabels(reviewer)).toEqual(expected);
    expect(textLabels(reviewer)).toEqual(expected);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(contentLabels(reopened)).toEqual(expected);
  });

  test("an item inserted with its own style keeps the list's number, as the save does", async () => {
    const reviewer = await reviewerOf(LIST);
    apply(reviewer, "direct", [
      {
        id: "1",
        type: "insertAfterBlock",
        blockId: blockId(reviewer, "Deposit on signature"),
        text: "Interim payment",
        styleId: "Quote",
      },
    ]);

    await expectLabels(reviewer, [
      "Payment happens in stages.",
      "1. Deposit on signature",
      "2. Interim payment",
      "3. Balance on delivery",
      "4. Retention after inspection",
      "Closing remarks.",
    ]);
  });
});

describe("a paragraph numbered at a level its list does not define", () => {
  test("shows no marker, reads as a paragraph that keeps its level, and does not count", async () => {
    const reviewer = await reviewerOf(LIST);
    const numId = reviewer.getContent().find(({ text }) => text === "Deposit on signature")
      ?.listReference?.numId;
    expect(numId).toBeDefined();
    apply(reviewer, "direct", [
      {
        id: "1",
        type: "insertAfterBlock",
        blockId: blockId(reviewer, "Deposit on signature"),
        text: "Level eight.",
        numbering: { numId: numId ?? 0, level: 8 },
      },
    ]);

    await expectLabels(reviewer, [
      "Payment happens in stages.",
      "1. Deposit on signature",
      "Level eight.",
      "2. Balance on delivery",
      "3. Retention after inspection",
      "Closing remarks.",
    ]);
    for (const current of [
      reviewer,
      await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()),
    ]) {
      expect(current.getContent().find(({ text }) => text === "Level eight.")).toMatchObject({
        kind: "paragraph",
        listLevel: 8,
        listReference: { numId, level: 8 },
      });
    }
  });
});
