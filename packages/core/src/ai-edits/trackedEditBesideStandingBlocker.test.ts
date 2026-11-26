/**
 * A tracked edit proves that Reject All gives the document before it back.
 * A revision the document already held that blocks its own resolution, the
 * same way before and after the edit, stays pending on both sides of that
 * proof: it is not the edit's, so it must not refuse every tracked edit of a
 * document that holds one, while the suggested mode admits the same edit.
 */

import { expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

// Rejecting the inserted mark of 00000002 would run its words into the table
// across the bookmark standing before it, so Reject All refuses that mark.
const document: Document = {
  package: {
    document: {
      content: [
        { type: "paragraph", paraId: "00000001", content: [] },
        {
          type: "paragraph",
          paraId: "00000002",
          content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
          pPrMark: { kind: "ins", info: { id: 2, author: "A" } },
        },
        { type: "bookmarkStart", id: 90, name: "_GoBack" },
        {
          type: "table",
          columnWidths: [2400, 2400],
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content: [
                    { type: "paragraph", paraId: "00000003", content: [] },
                  ],
                },
              ],
            },
          ],
        },
      ],
      finalSectionProperties: { pageWidth: 12240, pageHeight: 15840 },
    },
  },
};

const open = async () => {
  const prepared = await prepareCanonicalDocxInput(
    await createDocx(structuredClone(document)),
  );
  if (prepared.isErr()) throw prepared.error;
  return FolioDocxReviewer.fromBuffer(prepared.value.slice().buffer, {
    author: "Reviewer",
    session: "canonical",
  });
};

const insertion: FolioDocumentOperation = {
  id: "insert",
  type: "insertBeforeBlock",
  blockId: "00000001",
  text: "revised",
};

test("the standing blocker refuses Reject All before any edit", async () => {
  const reviewer = await open();
  expect(() => reviewer.rejectAll()).toThrow(/range boundary/u);
});

for (const mode of ["tracked-changes", "suggested"] as const)
  test(`a ${mode} paragraph insertion beside a standing blocker applies`, async () => {
    const reviewer = await open();
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations: [insertion],
        mode,
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    expect(
      reviewer
        .getChanges()
        .filter(({ author }) => author === "Reviewer")
        .map(({ type, text }) => ({ type, text })),
    ).toEqual([
      { type: "paragraphMarkInserted", text: "revised" },
      { type: "insertion", text: "revised" },
    ]);
  });
