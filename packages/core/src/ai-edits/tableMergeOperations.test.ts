/**
 * Row and column operations beside a vertical merge, through the public
 * reviewer: applied directly, applied tracked and then accepted or rejected,
 * each saved and reopened, and every table read back three ways.
 */

import { describe, expect, test } from "bun:test";

import {
  buildTableDocx,
  readReviewerTables,
  tableReadingProblems,
  type TableReading,
  type TableSpec,
} from "../__tests__/tableOperationDocument";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";

/** Two columns, three rows, `A1` merged down over the second row. */
const MERGED_LEFT: TableSpec = {
  rows: 3,
  columns: 2,
  cells: [
    { row: 0, column: 0, rowSpan: 2, columnSpan: 1, text: "A1" },
    { row: 0, column: 1, rowSpan: 1, columnSpan: 1, text: "B1" },
    { row: 1, column: 1, rowSpan: 1, columnSpan: 1, text: "B2" },
    { row: 2, column: 0, rowSpan: 1, columnSpan: 1, text: "A3" },
    { row: 2, column: 1, rowSpan: 1, columnSpan: 1, text: "B3" },
  ],
};

type Mode = "direct" | "tracked-changes";

const open = (bytes: ArrayBuffer) => FolioDocxReviewer.fromBuffer(bytes, { author: "Tester" });

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) {
    throw new Error(`no block reads ${JSON.stringify(text)}`);
  }
  return block.id;
};

const apply = async (
  base: ArrayBuffer,
  mode: Mode,
  build: (reviewer: FolioDocxReviewer) => Omit<FolioDocumentOperation, "id">,
) => {
  const reviewer = await open(base);
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [{ id: "op", ...build(reviewer) } as FolioDocumentOperation],
  });
  return { reviewer, result };
};

/** A tracked result, saved, reopened and resolved one way. */
const resolveTracked = async (
  reviewer: FolioDocxReviewer,
  resolution: "accept" | "reject",
): Promise<TableReading> => {
  const reopened = await open(await reviewer.toBuffer());
  expect(resolution === "accept" ? reopened.acceptAll() : reopened.rejectAll()).toBeGreaterThan(0);
  return readReviewerTables(reopened);
};

const cells = (reading: TableReading) =>
  reading.snapshot.map((table) =>
    table.cells.map(({ row, column, rowSpan, columnSpan, text }) => [
      `${row}:${column}`,
      `${rowSpan}x${columnSpan}`,
      text,
    ]),
  );

describe("inserting a row inside a vertical merge", () => {
  test.each(["direct", "tracked-changes"] as const)(
    "refuses a cell text the new row has no cell for, and changes nothing (%s)",
    async (mode) => {
      const base = await buildTableDocx(MERGED_LEFT);
      const original = await readReviewerTables(await open(base));
      const { reviewer, result } = await apply(base, mode, (live) => ({
        type: "insertTableRow",
        blockId: blockId(live, "B1"),
        position: "after",
        cellTexts: ["X", "Y"],
      }));

      expect(result.applied).toEqual([]);
      expect(result.issues).toEqual([
        expect.objectContaining({
          operationId: "op",
          code: "payloadDoesNotFit",
          recovery: "changeTarget",
          message: expect.stringContaining('cellTexts[1] "Y"'),
        }),
      ]);
      const after = await readReviewerTables(reviewer);
      expect(cells(after)).toEqual(cells(original));
    },
  );

  test("places every text it has a cell for, direct and tracked alike", async () => {
    const base = await buildTableDocx(MERGED_LEFT);
    const operation = (live: FolioDocxReviewer) =>
      ({
        type: "insertTableRow",
        blockId: blockId(live, "B1"),
        position: "after",
        cellTexts: ["X"],
      }) as const;
    const direct = await apply(base, "direct", operation);
    const tracked = await apply(base, "tracked-changes", operation);
    expect(direct.result.issues).toEqual([]);
    expect(tracked.result.issues).toEqual([]);

    const directReading = await readReviewerTables(direct.reviewer);
    const accepted = await resolveTracked(tracked.reviewer, "accept");
    const rejected = await resolveTracked(tracked.reviewer, "reject");
    const original = await readReviewerTables(await open(base));

    expect(cells(directReading)).toEqual([
      [
        ["0:0", "3x1", "A1"],
        ["0:1", "1x1", "B1"],
        ["1:1", "1x1", "X"],
        ["2:1", "1x1", "B2"],
        ["3:0", "1x1", "A3"],
        ["3:1", "1x1", "B3"],
      ],
    ]);
    expect(cells(accepted)).toEqual(cells(directReading));
    expect(cells(rejected)).toEqual(cells(original));
    for (const reading of [directReading, accepted, rejected]) {
      expect(tableReadingProblems(reading)).toEqual([]);
    }
  });
});
describe("deleting the column beside a vertical merge", () => {
  test("direct and accepted tracked deletion leave the same coherent table", async () => {
    const base = await buildTableDocx(MERGED_LEFT);
    const operation = (live: FolioDocxReviewer) =>
      ({ type: "deleteTableColumn", blockId: blockId(live, "B1") }) as const;
    const direct = await apply(base, "direct", operation);
    const tracked = await apply(base, "tracked-changes", operation);
    expect(direct.result.issues).toEqual([]);
    expect(tracked.result.issues).toEqual([]);

    const directReading = await readReviewerTables(direct.reviewer);
    const accepted = await resolveTracked(tracked.reviewer, "accept");
    const rejected = await resolveTracked(tracked.reviewer, "reject");
    const original = await readReviewerTables(await open(base));

    // The merged cell's second row held nothing but the merge once column B
    // went, so the row goes with it and the merge closes over one row.
    expect(cells(directReading)).toEqual([
      [
        ["0:0", "1x1", "A1"],
        ["1:0", "1x1", "A3"],
      ],
    ]);
    expect(cells(accepted)).toEqual(cells(directReading));
    expect(cells(rejected)).toEqual(cells(original));
    for (const reading of [directReading, accepted, rejected]) {
      expect(tableReadingProblems(reading)).toEqual([]);
    }
  });
});

describe("a column deletion in a nested table the same batch moves", () => {
  /**
   * The outer table merges whole, carrying the nested table into the merged
   * cell, while the nested table loses the only column its second row had a
   * cell in. The emptied row still has to close once the batch is done.
   */
  const NESTED: TableSpec = {
    rows: 3,
    columns: 4,
    cells: [
      { row: 0, column: 0, rowSpan: 1, columnSpan: 1, text: "" },
      {
        row: 0,
        column: 1,
        rowSpan: 2,
        columnSpan: 2,
        text: "c1",
        nested: {
          rows: 2,
          columns: 3,
          cells: [
            { row: 0, column: 0, rowSpan: 2, columnSpan: 2, text: "n0" },
            { row: 0, column: 2, rowSpan: 1, columnSpan: 1, text: "n1" },
            { row: 1, column: 2, rowSpan: 1, columnSpan: 1, text: "n2" },
          ],
        },
      },
      { row: 0, column: 3, rowSpan: 1, columnSpan: 1, text: "" },
      { row: 1, column: 0, rowSpan: 1, columnSpan: 1, text: "" },
      { row: 1, column: 3, rowSpan: 2, columnSpan: 1, text: "c4" },
      { row: 2, column: 0, rowSpan: 1, columnSpan: 2, text: "" },
      { row: 2, column: 2, rowSpan: 1, columnSpan: 1, text: "" },
    ],
  };

  test("closes the row the deletion emptied, in the editor and the package alike", async () => {
    const reviewer = await open(await buildTableDocx(NESTED));
    const firstCell = reviewer
      .getContent()
      .find(
        ({ table }) =>
          table?.tableIndex === 0 &&
          table.rowIndex === 0 &&
          table.gridColumnIndex === 0 &&
          table.paragraphIndex === 0,
      );
    if (!firstCell) {
      throw new Error("no block starts the first cell");
    }
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        {
          id: "merge",
          type: "mergeTableCells",
          blockId: blockId(reviewer, "c4"),
          endBlockId: firstCell.id,
        },
        { id: "column", type: "deleteTableColumn", blockId: blockId(reviewer, "n1") },
      ],
    });
    expect(result.applied.map(({ id }) => id).toSorted()).toEqual(["column", "merge"]);

    const reading = await readReviewerTables(reviewer);
    expect(tableReadingProblems(reading)).toEqual([]);
    expect(cells(reading)[1]).toEqual([["0:0", "1x2", "n0"]]);
  });
});
