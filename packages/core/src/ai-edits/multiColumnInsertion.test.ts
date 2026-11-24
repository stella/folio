/**
 * `insertTableColumn` with a `columnCount` inserts that many adjacent columns
 * as one change: tracked or suggested, accepting it gives the direct edit and
 * rejecting it gives the original, every column at once. On the canonical
 * session a second, separate tracked column insertion beside the pending one
 * is refused: a table records one pending grid change.
 */

import { describe, expect, test } from "bun:test";

import {
  buildTableDocx,
  readReviewerTables,
  type TableSpec,
} from "../__tests__/tableOperationDocument";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  parseFolioDocumentOperationBatch,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";

const GRID: TableSpec = {
  rows: 2,
  columns: 2,
  cells: ["Item", "Price", "Term", "Value"].map((text, slot) => ({
    row: Math.floor(slot / 2),
    column: slot % 2,
    rowSpan: 1,
    columnSpan: 1,
    text,
  })),
};

const SESSIONS = ["canonical"] as const;
const MODES = ["tracked-changes", "suggested"] as const;
const POSITIONS = ["before", "after"] as const;

describe.each(SESSIONS)("%s session", (session) => {
  const open = (bytes: ArrayBuffer) =>
    FolioDocxReviewer.fromBuffer(bytes, { author: "Tester", session });
  const blockId = (reviewer: FolioDocxReviewer, text: string) =>
    reviewer.getContent().find((block) => block.text === text)?.id ??
    `missing ${text}`;
  const insertion = (
    reviewer: FolioDocxReviewer,
    position: (typeof POSITIONS)[number],
  ): FolioDocumentOperation => ({
    id: "columns",
    type: "insertTableColumn",
    blockId: blockId(reviewer, "Value"),
    position,
    columnCount: 3,
    cellTexts: ["a1", "a2", "a3", "b1", "b2"],
  });
  const apply = (
    reviewer: FolioDocxReviewer,
    mode: "direct" | (typeof MODES)[number],
    operation: FolioDocumentOperation,
  ) =>
    reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [operation],
      },
      { undefinedReferences: "refuse" },
    );
  const cells = async (reviewer: FolioDocxReviewer) =>
    (await readReviewerTables(reviewer)).snapshot;

  test.each(POSITIONS)("direct: three columns %s a cell", async (position) => {
    const reviewer = await open(await buildTableDocx(GRID));
    expect(
      apply(reviewer, "direct", insertion(reviewer, position)).skipped,
    ).toEqual([]);
    const texts = (await cells(reviewer))
      .at(0)
      ?.cells.map(({ row, column, text }) => [row, column, text]);
    const first = position === "after" ? 2 : 1;
    expect(texts).toEqual(
      expect.arrayContaining([
        [0, first, "a1"],
        [0, first + 1, "a2"],
        [0, first + 2, "a3"],
        [1, first, "b1"],
        [1, first + 1, "b2"],
        [1, first + 2, ""],
      ]),
    );
    expect(texts).toHaveLength(10);
  });

  describe.each(MODES)("%s", (mode) => {
    test.each(POSITIONS)(
      "three columns %s a cell resolve as one change",
      async (position) => {
        const base = await buildTableDocx(GRID);
        const original = await cells(await open(base));
        const direct = await open(base);
        apply(direct, "direct", insertion(direct, position));
        const expected = await cells(direct);
        for (const decision of ["accept", "reject"] as const) {
          const reviewer = await open(base);
          const result = apply(reviewer, mode, insertion(reviewer, position));
          expect(result.skipped).toEqual([]);
          if (mode === "suggested") {
            expect(
              decision === "accept"
                ? reviewer.acceptSuggestion("columns")
                : reviewer.rejectSuggestion("columns"),
            ).toBe(true);
            if (decision === "accept") reviewer.acceptAll();
          } else if (decision === "accept") reviewer.acceptAll();
          else reviewer.rejectAll();
          expect(await cells(reviewer)).toEqual(
            decision === "accept" ? expected : original,
          );
        }
      },
    );
  });

  if (session === "canonical")
    test("a second tracked column insertion beside a pending one is refused", async () => {
      const reviewer = await open(await buildTableDocx(GRID));
      expect(
        apply(reviewer, "tracked-changes", insertion(reviewer, "after"))
          .skipped,
      ).toEqual([]);
      const second = apply(reviewer, "tracked-changes", {
        id: "second",
        type: "insertTableColumn",
        blockId: blockId(reviewer, "Item"),
      });
      expect(second.applied).toEqual([]);
      expect(second.skipped).toEqual([
        expect.objectContaining({
          id: "second",
          message: expect.stringContaining("one change in one operation"),
        }),
      ]);
    });
});

test.each([0, 64, 1.5])(
  "columnCount %p is not a valid batch",
  (columnCount) => {
    expect(() =>
      parseFolioDocumentOperationBatch({
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          {
            id: "columns",
            type: "insertTableColumn",
            blockId: "00000001",
            columnCount,
          },
        ],
      }),
    ).toThrow("columnCount");
  },
);
