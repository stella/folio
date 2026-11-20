/**
 * A tracked edit is checked by rejecting every change before and after it.
 * Rejecting a pending split of a list paragraph rebuilds that paragraph on
 * both sides; both read its indentation as the list's, so a merge elsewhere
 * in the story applies.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

const apply = (
  reviewer: FolioDocxReviewer,
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations,
    },
    { undefinedReferences: "refuse" },
  );

test("a tracked merge applies beside a pending split of a list paragraph", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(
      fromMarkdown(
        "- Goods are packed securely\n- Goods are insured in transit\n- Balance on delivery\n- Retention after inspection",
      ),
    ),
    { author: "Editor", session: "canonical" },
  );
  const [first, , third] = reviewer.getContent();
  if (first === undefined || third === undefined)
    throw new Error("The fixture has four list paragraphs.");
  expect(
    apply(reviewer, [
      { id: "split", type: "splitBlock", blockId: third.id, offset: 11 },
    ]).skipped,
  ).toEqual([]);
  const merged = apply(reviewer, [
    {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: first.id,
      separator: " ",
    },
  ]);
  expect(merged.skipped).toEqual([]);
  expect(merged.applied.map(({ id }) => id)).toEqual(["merge"]);
});
