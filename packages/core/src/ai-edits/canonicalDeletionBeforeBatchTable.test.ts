import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const anchor = "10000000";
const OPERATIONS: readonly FolioDocumentOperation[] = [
  { id: "table", type: "insertTable", blockId: anchor, rows: [["Term"]] },
  { id: "before", type: "insertBeforeBlock", blockId: anchor, text: "Before." },
  { id: "delete", type: "deleteBlock", blockId: anchor },
];
const PERMUTATIONS = [
  [0, 1, 2],
  [0, 2, 1],
  [1, 0, 2],
  [1, 2, 0],
  [2, 0, 1],
  [2, 1, 0],
];

const tracked = (
  reviewer: FolioDocxReviewer,
  operations: readonly FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [...operations],
    },
    { undefinedReferences: "refuse" },
  );

/** A paragraph whose own mark a tracked split inserted, holding original words. */
const splitDocument = async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await paragraphsDocx([[textRun("Alpha bravo.")]]),
    { session: "canonical" },
  );
  tracked(reviewer, [
    { id: "split", type: "splitBlock", blockId: anchor, offset: 5 },
  ]);
  return reviewer;
};

// Before a table the batch puts after it, the paragraph's deletion takes its
// words only, as it does once that table stands: tracked, in any order.
test("a deletion before a table its batch inserts reads as after that table stands", async () => {
  const sequential = await splitDocument();
  for (const operation of OPERATIONS)
    expect(tracked(sequential, [operation]).skipped).toEqual([]);
  sequential.acceptAll();
  const expected = sequential.getContent().map(({ text }) => text);
  for (const order of PERMUTATIONS) {
    const reviewer = await splitDocument();
    const result = tracked(
      reviewer,
      order.flatMap((index) => OPERATIONS[index] ?? []),
    );
    expect(result.skipped).toEqual([]);
    reviewer.acceptAll();
    expect(reviewer.getContent().map(({ text }) => text)).toEqual(expected);
  }
});
