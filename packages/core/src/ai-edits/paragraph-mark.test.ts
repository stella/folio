/**
 * `splitBlock` and `mergeBlockWithNext`: a paragraph mark moves and no words
 * do.
 *
 * The alternative the engine used before — rewrite the first half, insert the
 * second — round-trips and lies: it claims the tail was newly written when
 * nobody touched it. These operations put the change where it happened, on
 * the paragraph mark, which is also what `w:rPr/w:ins` inside `w:pPr` means
 * to any OOXML consumer.
 */

import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "./headless";
import { buildParagraphsDocx } from "./__fixtures__/paragraphs";

const STAMP = { date: "2024-03-01T00:00:00.000Z", idSeed: 500 };

const textsOf = async (buffer: ArrayBuffer, view: "original" | "final"): Promise<string[]> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer);
  return (reviewer.readReviewedStory({ view })?.snapshot.blocks ?? []).map(({ text }) => text);
};

type RunOptions = {
  paragraphs: readonly string[];
  operation: Parameters<FolioDocxReviewer["applyOperations"]>[0][number];
  mode?: "direct" | "tracked-changes";
  /** Index of the block the operation addresses. */
  blockIndex: number;
  /** Put the paragraphs in one table cell, with a body paragraph after it. */
  inTableCell?: boolean;
};

/** Apply one operation to a generated document and hand back its two views. */
const run = async ({ paragraphs, operation, mode, blockIndex, inTableCell }: RunOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await buildParagraphsDocx(paragraphs, { ...(inTableCell === true && { inTableCell }) }),
    { author: "compare" },
  );
  const snapshot = reviewer.snapshot();
  const blockId = snapshot.blocks[blockIndex]?.id ?? "";
  const { applied, skipped } = reviewer.applyOperations([{ ...operation, blockId }], {
    mode: mode ?? "tracked-changes",
    snapshot,
    revisionStamp: STAMP,
  });
  const buffer = await reviewer.toBuffer();
  return {
    applied,
    skipped,
    original: await textsOf(buffer, "original"),
    final: await textsOf(buffer, "final"),
  };
};

describe("splitBlock", () => {
  test("an inserted paragraph mark: accepting keeps the break, rejecting closes it", async () => {
    const { skipped, original, final } = await run({
      paragraphs: ["The buyer shall pay the seller.", "Delivery follows."],
      operation: { id: "split", type: "splitBlock", blockId: "", offset: 15, separator: " " },
      blockIndex: 0,
    });

    expect(skipped).toEqual([]);
    expect(original).toEqual(["The buyer shall pay the seller.", "Delivery follows."]);
    expect(final).toEqual(["The buyer shall", "pay the seller.", "Delivery follows."]);
  });

  test("without a separator the break keeps every character", async () => {
    const { skipped, original, final } = await run({
      paragraphs: ["The buyer shall pay the seller."],
      operation: { id: "split", type: "splitBlock", blockId: "", offset: 16 },
      blockIndex: 0,
    });

    expect(skipped).toEqual([]);
    expect(original).toEqual(["The buyer shall pay the seller."]);
    expect(final).toEqual(["The buyer shall ", "pay the seller."]);
  });

  test("direct mode splits without a revision", async () => {
    const { skipped, original, final } = await run({
      paragraphs: ["The buyer shall pay the seller."],
      operation: { id: "split", type: "splitBlock", blockId: "", offset: 15, separator: " " },
      mode: "direct",
      blockIndex: 0,
    });

    expect(skipped).toEqual([]);
    expect(original).toEqual(["The buyer shall", "pay the seller."]);
    expect(final).toEqual(["The buyer shall", "pay the seller."]);
  });

  test("refuses a separator that is not there any more", async () => {
    const { skipped } = await run({
      paragraphs: ["The buyer shall pay the seller."],
      operation: { id: "split", type: "splitBlock", blockId: "", offset: 4, separator: " " },
      blockIndex: 0,
    });

    expect(skipped).toEqual([{ id: "split", reason: "staleRange" }]);
  });

  test("refuses a split at either end, which is an insertion rather than a break", async () => {
    for (const offset of [0, 31]) {
      const { skipped } = await run({
        paragraphs: ["The buyer shall pay the seller."],
        operation: { id: "split", type: "splitBlock", blockId: "", offset },
        blockIndex: 0,
      });
      expect(skipped).toEqual([{ id: "split", reason: "staleRange" }]);
    }
  });
});

describe("mergeBlockWithNext", () => {
  test("a deleted paragraph mark: accepting closes the break, rejecting keeps it", async () => {
    const { skipped, original, final } = await run({
      paragraphs: ["The buyer shall", "pay the seller."],
      operation: { id: "merge", type: "mergeBlockWithNext", blockId: "", separator: " " },
      blockIndex: 0,
    });

    expect(skipped).toEqual([]);
    expect(original).toEqual(["The buyer shall", "pay the seller."]);
    expect(final).toEqual(["The buyer shall pay the seller."]);
  });

  test("direct mode joins without a revision", async () => {
    const { skipped, final } = await run({
      paragraphs: ["The buyer shall", "pay the seller."],
      operation: { id: "merge", type: "mergeBlockWithNext", blockId: "", separator: " " },
      mode: "direct",
      blockIndex: 0,
    });

    expect(skipped).toEqual([]);
    expect(final).toEqual(["The buyer shall pay the seller."]);
  });

  test("refuses the last paragraph of the story, which has nothing to join with", async () => {
    const { skipped } = await run({
      paragraphs: ["The buyer shall", "pay the seller."],
      operation: { id: "merge", type: "mergeBlockWithNext", blockId: "", separator: " " },
      blockIndex: 1,
    });

    expect(skipped).toEqual([{ id: "merge", reason: "unsupportedBlock" }]);
  });

  test("refuses the last paragraph of a table cell, which has a sibling but not one it can join", async () => {
    // A paragraph mark does not span a cell boundary. The paragraph after
    // this one is the body paragraph following the table, so a naive "is
    // there a next block" check would say yes and accepting the revision
    // would then have to do something no OOXML consumer can express.
    const { skipped } = await run({
      paragraphs: ["The buyer shall", "pay the seller."],
      operation: { id: "merge", type: "mergeBlockWithNext", blockId: "", separator: " " },
      blockIndex: 1,
      inTableCell: true,
    });

    expect(skipped).toEqual([{ id: "merge", reason: "unsupportedBlock" }]);
  });

  test("joins inside a table cell when there is a sibling paragraph in it", async () => {
    const { skipped, original, final } = await run({
      paragraphs: ["The buyer shall", "pay the seller."],
      operation: { id: "merge", type: "mergeBlockWithNext", blockId: "", separator: " " },
      blockIndex: 0,
      inTableCell: true,
    });

    expect(skipped).toEqual([]);
    expect(original).toEqual(["The buyer shall", "pay the seller.", "After the table."]);
    expect(final).toEqual(["The buyer shall pay the seller.", "After the table."]);
  });
});
