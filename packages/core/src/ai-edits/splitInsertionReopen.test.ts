/**
 * An edit inside a pending insertion leaves the insertion in stretches with
 * other content between them. The save writes each stretch as a revision of
 * its own, so the reviewer lists one change per stretch before the save too:
 * the list reads the same once reopened.
 */

import { describe, expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const INSERTED = "Agreement supplier buyer.";

const changes = (reviewer: FolioDocxReviewer) =>
  reviewer.getChanges().map(({ type, author, text, blockId }) => ({ type, author, text, blockId }));

describe("an edit inside a pending insertion across a save", () => {
  test.each([
    ["leading", "Agreement", "updated", undefined],
    ["middle", "supplier", "vendor", undefined],
    ["leading", "Agreement", "updated", "Heading2"],
  ] as const)(
    "replacing its %s words %s with %s (style %s)",
    async (_label, find, replace, styleId) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await createDocx(fromMarkdown("# Title\n\nFirst clause.\n\nLast clause.")),
        { author: "AI" },
      );
      const apply = (operation: Record<string, unknown>) =>
        expect(
          reviewer.applyDocumentOperations({
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            mode: "tracked-changes",
            operations: [{ id: String(operation["type"]), ...operation }],
          } as never).skipped,
        ).toEqual([]);
      const first = reviewer.getContent().find(({ text }) => text === "First clause.")?.id;
      apply({
        type: "insertAfterBlock",
        blockId: first,
        text: INSERTED,
        ...(styleId && { styleId }),
      });
      const inserted = reviewer.getContent().find(({ text }) => text === INSERTED)?.id;
      apply({ type: "replaceInBlock", blockId: inserted, find, replace });

      const before = changes(reviewer);
      const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
      expect(changes(reopened)).toEqual(before);
    },
  );
});
