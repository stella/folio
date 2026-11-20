/**
 * A direct deletion of a paragraph no paragraph follows joins the paragraph
 * before it as the batch stands when the deletion runs: when the same batch
 * first inserts a paragraph right before it, that paragraph is the one it
 * joins.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

test("a deletion of the last paragraph joins a paragraph the batch inserts before it", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(
      fromMarkdown("Opening words.\n\nMiddle words.\n\nClosing words."),
    ),
    { author: "Editor", session: "canonical" },
  );
  const [, middle, closing] = reviewer.snapshot().blocks;
  if (middle === undefined || closing === undefined)
    throw new Error("The fixture has three paragraphs.");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      atomic: true,
      operations: [
        {
          id: "insert",
          type: "insertAfterBlock",
          blockId: middle.id,
          text: "Added words.",
        },
        { id: "delete", type: "deleteBlock", blockId: closing.id },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied.map(({ id }) => id).sort()).toEqual([
    "delete",
    "insert",
  ]);
  const blocks = reviewer.snapshot().blocks;
  expect(blocks.map(({ text }) => text)).toEqual([
    "Opening words.",
    "Middle words.",
    "Added words.",
  ]);
  // The receipt names the paragraph the join retired: the inserted one, not
  // the paragraph before it, which stays.
  const deleted = result.applied.find(({ id }) => id === "delete");
  const identity = deleted?.paragraphIdentity;
  if (identity?.type !== "joined")
    throw new Error("The deletion reports the join it made.");
  expect(identity.blockId).toBe(closing.id);
  expect(blocks.map(({ id }) => id)).not.toContain(identity.retiredBlockId);
  expect(identity.retiredBlockId).not.toBe(middle.id);
});
