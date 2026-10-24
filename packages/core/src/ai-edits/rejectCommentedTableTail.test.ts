import { expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const SOURCE_TEXT = "Source paragraph.";
const ORIGINAL = [SOURCE_TEXT];

const open = async (bytes?: ArrayBuffer) => {
  if (bytes) return FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });
  const document = fromMarkdown(SOURCE_TEXT);
  const { docx } = await ensureParaIds(
    new Uint8Array(await createDocx(document)),
  );
  return FolioDocxReviewer.fromBuffer(new Uint8Array(docx).buffer, {
    author: "Reviewer",
  });
};

const visibleBlocks = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ text }) => text);

const applyTracked = (
  reviewer: FolioDocxReviewer,
  operations: readonly FolioDocumentOperation[],
): void => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations,
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(operations.length);
};

test("rejecting a comment on a paragraph inserted before a table restores the source", async () => {
  const reviewer = await open();
  const sourceId = reviewer.getContent().at(0)?.id;
  if (!sourceId) throw new Error("Missing source paragraph id");

  applyTracked(reviewer, [
    {
      id: "table",
      type: "insertTable",
      blockId: sourceId,
      rows: [["Term", "Value"]],
    },
  ]);
  applyTracked(reviewer, [
    {
      id: "paragraph",
      type: "insertAfterBlock",
      blockId: sourceId,
      text: "Inserted clause.",
    },
  ]);

  const insertedId = reviewer
    .getContent()
    .find(({ text }) => text === "Inserted clause.")?.id;
  if (!insertedId) throw new Error("Missing inserted paragraph id");
  applyTracked(reviewer, [
    {
      id: "comment",
      type: "commentOnBlock",
      blockId: insertedId,
      quote: "Inserted",
      comment: { text: "Comment on inserted text." },
    },
  ]);

  reviewer.rejectAll();
  const actual = visibleBlocks(reviewer);
  expect(actual).toEqual(ORIGINAL);
  expect(visibleBlocks(await open(await reviewer.toBuffer()))).toEqual(
    ORIGINAL,
  );
});
