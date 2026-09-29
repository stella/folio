/**
 * A tracked column edit marks one cell per row. The save writes each cell's
 * marker as a revision of its own, so the batch gives each cell its own id:
 * the reviewer lists, resolves, saves and reopens the same changes.
 */

import { describe, expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const MARKDOWN = [
  "# Schedule",
  "| Item | Price |\n|---|---|\n| Term | Value |\n| Period | 12 months |",
  "Closing.",
].join("\n\n");

const TYPES = ["insertTableColumn", "deleteTableColumn"] as const;

const isCellChange = ({ type }: { type: string }) =>
  type === "cellInserted" || type === "cellDeleted";

const cellChanges = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getChanges()
    .filter(isCellChange)
    .map(({ type, author, text, blockId }) => ({ type, author, text, blockId }));

const build = async (type: (typeof TYPES)[number]): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(fromMarkdown(MARKDOWN)), {
    author: "AI",
  });
  const blockId = reviewer.getContent().find((block) => block.text === "Value")?.id;
  if (blockId === undefined) throw new Error("fixture must expose the Value cell");
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [
      type === "insertTableColumn"
        ? { id: "op", type, blockId, position: "after" }
        : { id: "op", type, blockId },
    ],
  });
  expect(result.skipped).toEqual([]);
  expect(result.applied[0]?.revisionIds).toHaveLength(3);
  return reviewer;
};

const reopen = async (reviewer: FolioDocxReviewer): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());

describe("tracked column changes", () => {
  test.each(TYPES)("%s lists one change per cell, the same after a save", async (type) => {
    const reviewer = await build(type);
    const before = cellChanges(reviewer);
    expect(before).toHaveLength(3);
    const reopened = await reopen(reviewer);
    expect(cellChanges(reopened)).toEqual(before);
    for (const current of [reviewer, reopened]) {
      const ids = current
        .getChanges()
        .filter(isCellChange)
        .map(({ id }) => id);
      expect(new Set(ids).size).toBe(ids.length);
    }
  });

  test.each(TYPES)("%s: accepting one cell leaves the others pending", async (type) => {
    const reviewer = await build(type);
    const saved = await reopen(reviewer);
    for (const current of [reviewer, saved]) {
      const first = current.getChanges().find(isCellChange);
      if (!first) throw new Error("no cell change");
      expect(current.acceptChange(first)).toBe(true);
      expect(cellChanges(current)).toHaveLength(2);
    }
    expect(cellChanges(saved)).toEqual(cellChanges(reviewer));
  });
});
