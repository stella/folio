/**
 * A table that stops a pending join holds the joined words apart from the
 * paragraph after it. Removing that table would run those words into that
 * paragraph once accepted, which no operation asked for, so the removal is
 * refused until the join is resolved.
 */

import { describe, expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

type Setup = { pendingJoin: boolean };

/** "Alpha." merged into "Bravo." and "Bravo." deleted, tracked, before a table and "Charlie.". */
const beforeTable = async ({ pendingJoin }: Setup) => {
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
  const edits: FolioDocumentOperation[] = pendingJoin
    ? [
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: alpha.id,
          separator: " ",
        },
        { id: "delete", type: "deleteBlock", blockId: bravo.id },
      ]
    : [{ id: "delete", type: "deleteBlock", blockId: bravo.id }];
  for (const operation of edits) {
    expect(
      reviewer.applyOperations([operation], { mode: "tracked-changes" })
        .skipped,
    ).toEqual([]);
  }
  return { reviewer, cell: cell.id };
};

const deleteTable = (blockId: string): FolioDocumentOperation => ({
  id: "drop",
  type: "deleteTable",
  blockId,
});

describe("removing the table that stops a pending join", () => {
  for (const mode of ["direct", "tracked-changes", "suggested"] as const) {
    test(`is refused and changes nothing (${mode})`, async () => {
      const { reviewer, cell } = await beforeTable({ pendingJoin: true });
      const before = reviewer.snapshot();
      const result = reviewer.applyOperations([deleteTable(cell)], { mode });
      expect(result.applied).toEqual([]);
      expect(result.skipped.map(({ reason }) => reason)).toEqual([
        "pendingParagraphMarkDeletion",
      ]);
      expect(reviewer.snapshot()).toEqual(before);
    });
  }

  test("is applied once the join is rejected, and keeps the paragraphs apart", async () => {
    const { reviewer, cell } = await beforeTable({ pendingJoin: true });
    const join = reviewer
      .getChanges()
      .find(({ type }) => type === "paragraphMarkDeleted");
    if (!join) throw new Error("Missing join");
    expect(reviewer.rejectChange(join)).toBe(true);
    expect(
      reviewer.applyOperations([deleteTable(cell)], {
        mode: "tracked-changes",
      }).skipped,
    ).toEqual([]);
    reviewer.acceptAll();
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual([
      "Alpha.",
      "Charlie.",
    ]);
  });

  for (const mode of ["direct", "tracked-changes"] as const) {
    test(`is applied where no join runs into the table (${mode})`, async () => {
      const { reviewer, cell } = await beforeTable({ pendingJoin: false });
      expect(
        reviewer.applyOperations([deleteTable(cell)], { mode }).skipped,
      ).toEqual([]);
      reviewer.acceptAll();
      expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual([
        "Alpha.",
        "Charlie.",
      ]);
    });
  }
});
