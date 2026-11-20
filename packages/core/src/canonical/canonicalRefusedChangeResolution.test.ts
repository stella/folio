/**
 * Resolving one change the canonical engine will not resolve on its own
 * refuses it: `acceptChange` and `rejectChange` read `false` and leave the
 * document as it was, as for a change a locked content control holds.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

test("rejecting the break a table inserted after the last paragraph added is refused, not thrown", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(fromMarkdown("Opening words.\n\nClosing words.")),
    { author: "Editor", session: "canonical" },
  );
  const closing = reviewer.getContent().at(-1);
  if (closing === undefined) throw new Error("The fixture has paragraphs.");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "table",
          type: "insertTable",
          blockId: closing.id,
          rows: [["Term", "Value"]],
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  // The closing paragraph's break the insertion added: rejecting it alone
  // would leave the body ending with the table.
  const addedBreak = reviewer
    .getChanges()
    .find(
      ({ type, blockId }) =>
        type === "paragraphMarkInserted" && blockId === closing.id,
    );
  if (addedBreak === undefined) throw new Error("The insertion adds a break.");
  const before = reviewer.toDocument();
  expect(reviewer.rejectChange(addedBreak)).toBe(false);
  expect(reviewer.toDocument()).toEqual(before);
  expect(reviewer.getChanges().map(({ id }) => id)).toContain(addedBreak.id);
});
