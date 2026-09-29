/**
 * A tracked column edit marks one cell per row with the same revision. The
 * save gives each cell's marker an id of its own, so the reviewer lists one
 * change per cell before the save too: the list reads the same once reopened.
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

const cellChanges = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getChanges()
    .filter(({ type }) => type === "cellInserted" || type === "cellDeleted")
    .map(({ type, author, text, blockId }) => ({ type, author, text, blockId }));

describe("tracked column changes across a save", () => {
  test.each(["insertTableColumn", "deleteTableColumn"] as const)("%s", async (type) => {
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

    const before = cellChanges(reviewer);
    expect(before).toHaveLength(3);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(cellChanges(reopened)).toEqual(before);
  });
});
