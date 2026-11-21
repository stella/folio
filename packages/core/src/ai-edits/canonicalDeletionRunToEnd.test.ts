import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const TEXTS = ["P0.", "P1.", "P2.", "P3.", "P4.", "P5.", "P6."];
const idOf = (index: number) =>
  (0x1000_0000 + index).toString(16).toUpperCase();
const deletion = (index: number): FolioDocumentOperation => ({
  id: `delete-${String(index)}`,
  type: "deleteBlock",
  blockId: idOf(index),
});

const direct = async (operations: FolioDocumentOperation[]) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await paragraphsDocx(TEXTS.map((text) => [textRun(text)])),
    { session: "canonical" },
  );
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations,
    },
    { undefinedReferences: "refuse" },
  );
  return {
    skipped: result.skipped.map(({ id }) => id),
    texts: reviewer.getContent().map(({ text }) => text),
  };
};

// The last paragraph's deletion joins the paragraph before it as the batch
// runs: after the batch's deletions of the paragraphs right before it, the
// one that stays, in whatever order the batch lists them.
test("a direct batch deleting a run of paragraphs to the end removes them all", async () => {
  for (const run of [
    [5, 6],
    [6, 5],
    [4, 5, 6],
    [6, 4, 5],
  ])
    expect(await direct(run.map(deletion))).toEqual({
      skipped: [],
      texts: TEXTS.slice(0, Math.min(...run)),
    });
});

test("insertions elsewhere do not change which paragraph the last one joins", async () => {
  expect(
    await direct([
      deletion(6),
      {
        id: "insert-0",
        type: "insertBeforeBlock",
        blockId: idOf(0),
        text: "Start.",
      },
      deletion(5),
      {
        id: "insert-3",
        type: "insertBeforeBlock",
        blockId: idOf(3),
        text: "Middle.",
      },
    ]),
  ).toEqual({
    skipped: [],
    texts: ["Start.", "P0.", "P1.", "P2.", "Middle.", "P3.", "P4."],
  });
});
