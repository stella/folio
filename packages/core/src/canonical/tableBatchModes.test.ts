import { describe, expect, test } from "bun:test";

import { buildTableDocx } from "../__tests__/tableOperationDocument";
import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { parseDocx } from "../docx/parser";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";

const SESSIONS = [
  "canonical",
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
  // The merge claims both cells before the following row deletion can edit them.
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

      expect(result.applied.map(({ id }) => id)).toEqual(["merge"]);
      expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual([
        { id: "delete-row", reason: "overlappingOperation" },
      ]);
      if (mode === "tracked-changes") {
        const reopened = await FolioDocxReviewer.fromBuffer(
          await reviewer.toBuffer(),
          { session },
        );
        expect(reopened.acceptAll().resolved).toBeGreaterThan(0);
        const saved = await parseDocx(await reopened.toBuffer(), {
          preloadFonts: false,
        });
        const table = saved.package.document.content.find(
          (block) => block.type === "table",
        );
        expect(table?.rows.at(0)?.cells).toHaveLength(1);
        expect(table?.rows.at(0)?.cells.at(0)?.content).toHaveLength(2);
      }
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

  // A column inserted beside a merge down still moves the grid columns the
  // merge names, so the merge is refused whichever side the column lands on.
  test.each(["direct", "tracked-changes"] as const)(
    "a %s merge down beside a column the batch inserts is refused",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildTableDocx({
          rows: 2,
          columns: 1,
          cells: [
            { row: 0, column: 0, rowSpan: 1, columnSpan: 1, text: "top" },
            { row: 1, column: 0, rowSpan: 1, columnSpan: 1, text: "" },
          ],
        }),
        { session },
      );
      const blockId = cellBlockId(reviewer, 0);
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            { id: "merge", type: "mergeTableCells", blockId, rowCount: 2 },
            {
              id: "insert-column",
              type: "insertTableColumn",
              blockId,
              position: "before",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );

      expect(result.applied.map(({ id }) => id)).toEqual(["insert-column"]);
      expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual([
        { id: "merge", reason: "unsupportedBlock" },
      ]);
    },
  );

  // The row is named through the paragraph the batch deletes; the deletion
  // runs first, so the row is found through a paragraph the batch keeps.
  test.each(["direct", "tracked-changes"] as const)(
    "a %s row deletion named through a paragraph the batch deletes applies",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildTableDocx({
          rows: 1,
          columns: 1,
          cells: [
            {
              row: 0,
              column: 0,
              rowSpan: 1,
              columnSpan: 1,
              text: "c0",
              nested: {
                rows: 1,
                columns: 1,
                cells: [
                  { row: 0, column: 0, rowSpan: 1, columnSpan: 1, text: "n0" },
                ],
              },
            },
          ],
        }),
        { session },
      );
      const blockId = cellBlockId(reviewer, 0);
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            { id: "delete-paragraph", type: "deleteBlock", blockId },
            { id: "delete-row", type: "deleteTableRow", blockId },
          ],
        },
        { undefinedReferences: "refuse" },
      );

      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id).sort()).toEqual([
        "delete-paragraph",
        "delete-row",
      ]);
    },
  );

  // A column deletion is named through a paragraph outside the column; the
  // batch deletes that one, so another kept paragraph outside it names it.
  test.each(["direct", "tracked-changes"] as const)(
    "a %s column deletion named through a paragraph the batch deletes applies",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildTableDocx({
          rows: 1,
          columns: 2,
          cells: [
            { row: 0, column: 0, rowSpan: 1, columnSpan: 1, text: "c0" },
            {
              row: 0,
              column: 1,
              rowSpan: 1,
              columnSpan: 1,
              text: "c1",
              nested: {
                rows: 1,
                columns: 1,
                cells: [
                  { row: 0, column: 0, rowSpan: 1, columnSpan: 1, text: "n0" },
                ],
              },
            },
          ],
        }),
        { session },
      );
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "delete-paragraph",
              type: "deleteBlock",
              blockId: cellBlockId(reviewer, 1),
            },
            {
              id: "delete-column",
              type: "deleteTableColumn",
              blockId: cellBlockId(reviewer, 0),
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );

      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id).sort()).toEqual([
        "delete-column",
        "delete-paragraph",
      ]);
    },
  );

  // A bookmark around the whole table stands beside its rows, which no
  // column edit changes, so a suggested column edit keeps it where it is.
  test.each(["direct", "tracked-changes"] as const)(
    "a %s column edit on a table a bookmark wraps keeps the bookmark",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildTableDocx(ONE_ROW, (table) => ({
          ...table,
          bookmarks: [
            {
              index: 0,
              marker: { type: "bookmarkStart", id: 7, name: "Whole" },
            },
            { index: 1, marker: { type: "bookmarkEnd", id: 7 } },
          ],
        })),
        { session },
      );
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "insert-column",
              type: "insertTableColumn",
              blockId: cellBlockId(reviewer, 0),
              position: "after",
              cellTexts: ["new"],
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id)).toEqual(["insert-column"]);
      const saved = await parseDocx(await reviewer.toBuffer(), {
        preloadFonts: false,
        detectVariables: false,
      });
      const table = saved.package.document.content.find(
        (block) => block.type === "table",
      );
      if (table?.type !== "table") throw new TypeError("The table is gone.");
      expect(
        table.bookmarks?.map(({ index, marker }) => [marker.type, index]),
      ).toEqual([
        ["bookmarkStart", 0],
        ["bookmarkEnd", 1],
      ]);
      expect(table.rows.at(0)?.cells).toHaveLength(3);
    },
  );

  // A row content control wraps its row whatever cells it holds, so a
  // suggested column edit keeps the row inside it.
  test.each(["direct", "tracked-changes"] as const)(
    "a %s column edit in a row a content control wraps keeps the control",
    async (mode) => {
      const control = { sdtType: "richText" as const, id: 41, tag: "row" };
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await buildTableDocx(ONE_ROW, (table) => ({
          ...table,
          rows: table.rows.map((row) => ({
            ...row,
            contentControls: [control],
          })),
        })),
        { session },
      );
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "insert-column",
              type: "insertTableColumn",
              blockId: cellBlockId(reviewer, 0),
              position: "after",
              cellTexts: ["new"],
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id)).toEqual(["insert-column"]);
      const saved = await parseDocx(await reviewer.toBuffer(), {
        preloadFonts: false,
        detectVariables: false,
      });
      const table = saved.package.document.content.find(
        (block) => block.type === "table",
      );
      if (table?.type !== "table") throw new TypeError("The table is gone.");
      expect(
        table.rows.map((row) => row.contentControls?.map(({ tag }) => tag)),
      ).toEqual([["row"]]);
      expect(table.rows.at(0)?.cells).toHaveLength(3);
    },
  );
});
