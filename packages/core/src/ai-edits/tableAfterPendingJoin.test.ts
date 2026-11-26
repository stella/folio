/**
 * A table put after a paragraph never takes words once every change is
 * accepted. Accepting a pending deletion of a paragraph's break runs its
 * words into whatever follows, so a table put right after it is refused;
 * so is one put after a paragraph that goes whole, while a paragraph before
 * it joins into it: that paragraph's words run on through it into the table.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const IDS = ["0000000A", "0000000B", "0000000C", "0000000D"] as const;

const source = async (): Promise<ArrayBuffer> => {
  const document = fromMarkdown(
    "Terms\n\nClause one\n\nBody text.\n\nClause two",
  );
  document.package.document.content.forEach((block, index) => {
    const id = IDS[index];
    if (block.type === "paragraph" && id !== undefined) block.paraId = id;
  });
  const prepared = await prepareCanonicalDocxInput(await createDocx(document));
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

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

/** Tracked edits that leave breaks pending deletion before the table goes in. */
const SETUPS: Record<string, FolioDocumentOperation[]> = {
  // The last paragraph's deletion takes the break before the one it follows.
  "the last two paragraphs deleted": [
    { id: "c", type: "deleteBlock", blockId: "0000000C" },
    { id: "d", type: "deleteBlock", blockId: "0000000D" },
  ],
  "a middle paragraph deleted": [
    { id: "b", type: "deleteBlock", blockId: "0000000B" },
  ],
  "a merge, then the paragraph it joins deleted": [
    {
      id: "m",
      type: "mergeBlockWithNext",
      blockId: "0000000A",
      separator: " ",
    },
    { id: "b", type: "deleteBlock", blockId: "0000000B" },
  ],
};

describe("a tracked table after a paragraph beside a pending join", () => {
  for (const [setup, edits] of Object.entries(SETUPS))
    for (const anchor of IDS)
      test(`takes no words once accepted (${setup}, after ${anchor})`, async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
          author: "Reviewer",
          session: "canonical",
        });
        expect(apply(reviewer, edits).skipped).toEqual([]);
        const result = apply(reviewer, [
          {
            id: "table",
            type: "insertTable",
            blockId: anchor,
            rows: [
              ["Term", "Value"],
              ["Period", "12 months"],
            ],
          },
        ]);
        reviewer.acceptAll();
        const cells = reviewer
          .getContent()
          .filter(({ table }) => table !== undefined)
          .map(({ text }) => text);
        if (result.applied.length === 0) {
          expect(result.skipped).toMatchObject([
            { reason: "pendingParagraphMarkDeletion" },
          ]);
          expect(cells).toEqual([]);
        } else expect(cells).toEqual(["Term", "Value", "Period", "12 months"]);
      });
});
