/**
 * A paragraph inserted before an opaque block that ends the story (an
 * altChunk) goes whole when it is rejected: the block after it stays, so the
 * story ends as it did before the insertion. The saved package holds a
 * pending suggestion as rejected, so it must not keep an empty paragraph
 * there, and the suggestion restages on it as the suggesting reviewer shows.
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

const document: Document = {
  package: {
    document: {
      content: [
        { type: "paragraph", paraId: "00000001", content: [] },
        {
          type: "paragraph",
          paraId: "00000002",
          content: [],
          sectionProperties: { pageWidth: 11906 },
        },
        { type: "preservedBlock", xml: "<w:altChunk/>" },
      ],
      finalSectionProperties: { pageWidth: 12240, pageHeight: 15840 },
    },
  },
};

const source = async () => {
  const prepared = await prepareCanonicalDocxInput(
    await createDocx(structuredClone(document)),
  );
  if (prepared.isErr()) throw prepared.error;
  return prepared.value.slice().buffer;
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Reviewer",
    session: "canonical",
  });

const insertion: FolioDocumentOperation = {
  id: "insert",
  type: "insertAfterBlock",
  blockId: "00000002",
  text: "revised",
};

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "tracked-changes" | "suggested",
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [insertion],
      mode,
    },
    { undefinedReferences: "refuse" },
  );

const blocks = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ id, text }) => ({ id, text }));

test("a suggested paragraph before an ending altChunk saves as rejected and restages", async () => {
  const buffer = await source();
  const original = blocks(await open(buffer));
  const reviewer = await open(buffer);
  expect(apply(reviewer, "suggested").skipped).toEqual([]);
  const suggested = blocks(reviewer);
  const records = reviewer.exportPendingSuggestions();

  const saved = await reviewer.toBuffer();
  expect(blocks(await open(saved))).toEqual(original);

  const restaged = await open(saved);
  for (const loaded of restaged.loadPendingSuggestions(records))
    expect(loaded).toMatchObject({ status: "restaged" });
  expect(blocks(restaged)).toEqual(suggested);
});

test("a tracked paragraph before an ending altChunk rejects to the original", async () => {
  const buffer = await source();
  const original = blocks(await open(buffer));
  const reviewer = await open(buffer);
  expect(apply(reviewer, "tracked-changes").skipped).toEqual([]);
  reviewer.rejectAll();
  expect(blocks(reviewer)).toEqual(original);
});
