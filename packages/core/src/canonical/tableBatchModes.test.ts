import { describe, expect, test } from "bun:test";

import { buildTableDocx } from "../__tests__/tableOperationDocument";
import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";

const SESSIONS = [
  "canonical",
  "legacy",
] as const satisfies readonly HeadlessReviewerSession[];

const ONE_ROW = {
  rows: 1,
  columns: 2,
  cells: [
    { row: 0, column: 0, rowSpan: 1, columnSpan: 1, text: "c0" },
    { row: 0, column: 1, rowSpan: 1, columnSpan: 1, text: "c1" },
  ],
};

const cellBlockId = (reviewer: FolioDocxReviewer, column: number): string =>
  reviewer
    .snapshot()
    .blocks.find(
      ({ table }) => table?.rowIndex === 0 && table.gridColumnIndex === column,
    )?.id ?? "";

describe.each(SESSIONS)("%s session", (session) => {
  // A merge across a row moves content between cells, which no tracked
  // change records, so a tracked batch refuses it. It still claims the cells
  // it would merge: the row deletion after it is refused in either mode.
  test.each(["direct", "tracked-changes"] as const)(
    "a %s merge claims its cells against a later row deletion",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildTableDocx(ONE_ROW),
        { session },
      );
      const blockId = cellBlockId(reviewer, 0);
      const operations: FolioDocumentOperation[] = [
        {
          id: "merge",
          type: "mergeTableCells",
          blockId,
          endBlockId: cellBlockId(reviewer, 1),
        },
        { id: "delete-row", type: "deleteTableRow", blockId },
      ];
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations,
        },
        { undefinedReferences: "refuse" },
      );

      expect(result.applied.map(({ id }) => id)).toEqual(
        mode === "direct" ? ["merge"] : [],
      );
      expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual([
        ...(mode === "direct"
          ? []
          : [{ id: "merge", reason: "unsupportedBlock" }]),
        { id: "delete-row", reason: "overlappingOperation" },
      ]);
    },
  );

  // The column lands between the cells the merge names as the batch read
  // the table, so the merge would take the column's new cell with it.
  test.each(["direct", "tracked-changes"] as const)(
    "a %s merge across a column the batch inserts is refused",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildTableDocx(ONE_ROW),
        { session },
      );
      const blockId = cellBlockId(reviewer, 0);
      const operations: FolioDocumentOperation[] = [
        {
          id: "insert-column",
          type: "insertTableColumn",
          blockId,
          position: "after",
          cellTexts: ["new"],
        },
        {
          id: "merge",
          type: "mergeTableCells",
          blockId,
          endBlockId: cellBlockId(reviewer, 1),
        },
      ];
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations,
        },
        { undefinedReferences: "refuse" },
      );

      expect(result.applied.map(({ id }) => id)).toEqual(["insert-column"]);
      expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual([
        { id: "merge", reason: "unsupportedBlock" },
      ]);
    },
  );
});
