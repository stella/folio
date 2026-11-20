import { afterAll, beforeAll, expect, setSystemTime, test } from "bun:test";

import { paragraphsDocx, textRun } from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";

beforeAll(() => setSystemTime(new Date("2026-01-01T12:00:00Z")));
afterAll(() => setSystemTime());

const only = "10000000";
const remove: FolioDocumentOperation = {
  id: "remove",
  type: "deleteBlock",
  blockId: only,
};
const insert: FolioDocumentOperation = {
  id: "insert",
  type: "insertBeforeBlock",
  blockId: only,
  text: "Before.",
};

// The paragraph the batch puts before a container's only paragraph takes its
// break, so deleting that paragraph removes it whatever the input order.
for (const mode of ["direct", "tracked-changes"] as const)
  test(`${mode}: a sole paragraph deleted beside an insertion before it goes, in either input order`, async () => {
    for (const operations of [
      [remove, insert],
      [insert, remove],
    ]) {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await paragraphsDocx([[textRun("Only.")]]),
        { session: "canonical" },
      );
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations,
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
      reviewer.acceptAll();
      expect(reviewer.getContent().map(({ text }) => text)).toEqual([
        "Before.",
      ]);
    }
  });
