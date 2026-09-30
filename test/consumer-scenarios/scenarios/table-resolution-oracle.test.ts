import assert from "node:assert/strict";
import { describe, test } from "node:test";

import { compareWithModel, expectOperation, modelOf, type Row } from "../support/oracle.ts";
import { sequentialGroups } from "../support/metamorphic.ts";

type CellLocation = { rowIndex: number; column: number; columnSpan?: number; tableIndex?: number };

const row = (
  id: string,
  text: string,
  { rowIndex, column, columnSpan = 1, tableIndex = 0 }: CellLocation,
): Row => ({
  id,
  text,
  kind: "paragraph",
  table: {
    outerTableIndex: tableIndex,
    tableIndex,
    rowIndex,
    cellIndex: column,
    gridColumnIndex: column,
    columnSpan,
    rowSpan: 1,
  },
});

describe("table operation source coordinates", () => {
  test("replays row payloads before columns in source-grid execution order", () => {
    for (let width = 3; width <= 5; width += 1) {
      const rows = Array.from({ length: 2 }, (_, rowIndex) => {
        const cells: Row[] = [];
        for (let column = 0; column < width;) {
          const columnSpan = rowIndex === 0 && column === 0 ? 2 : 1;
          cells.push(row(`t0r${rowIndex}c${column}`, "cell", { rowIndex, column, columnSpan }));
          column += columnSpan;
        }
        return cells;
      }).flat();
      rows.push(
        ...Array.from({ length: width }, (_, column) =>
          row(`t1c${column}`, "cell", { rowIndex: 0, column, tableIndex: 1 }),
        ),
      );
      const operations = [
        { id: "delete-col-0", type: "deleteTableColumn", blockId: "t0r0c0" },
        { id: "insert-col-last", type: "insertTableColumn", blockId: `t0r0c${width - 1}` },
        {
          id: "row",
          type: "insertTableRow",
          blockId: "t0r1c0",
          cellTexts: Array.from({ length: width }, (_, index) => `cell ${index}`),
        },
        { id: "delete-col-2", type: "deleteTableColumn", blockId: "t0r0c2" },
        { id: "insert-col-tie", type: "insertTableColumn", blockId: "t0r0c0" },
        { id: "insert-other-table", type: "insertTableColumn", blockId: "t1c0", position: "after" },
      ];
      const ordered = sequentialGroups(operations, rows).flatMap((group) => group);
      assert.deepEqual(
        ordered.map(({ id }) => id),
        [
          "row",
          "insert-other-table",
          "insert-col-last",
          "insert-col-tie",
          "delete-col-2",
          "delete-col-0",
        ],
        `table width ${width}`,
      );
    }
  });

  test("places a row beside a pending deleted row in live row order", () => {
    for (let rowCount = 2; rowCount <= 5; rowCount += 1) {
      for (let deletedRow = 0; deletedRow < rowCount; deletedRow += 1) {
        const live = Array.from({ length: rowCount }, (_, rowIndex) => [
          row(`r${rowIndex}c0`, `Row ${rowIndex} cell 1`, { rowIndex, column: 0 }),
          row(`r${rowIndex}c1`, `Row ${rowIndex} cell 2`, { rowIndex, column: 1 }),
        ]).flat();
        const accepted = live
          .filter((cell) => cell.table?.rowIndex !== deletedRow)
          .map((cell) => {
            assert.ok(cell.table);
            return row(cell.id, cell.text, {
              rowIndex: cell.table.rowIndex - Number(cell.table.rowIndex > deletedRow),
              column: cell.table.gridColumnIndex,
            });
          });
        const pending = live.map((cell) => {
          assert.ok(cell.table);
          return cell.table.rowIndex === deletedRow ? Object.assign({}, cell, { text: "" }) : cell;
        });
        const model = modelOf(accepted, pending);
        model.mode = "tracked-changes";
        expectOperation(model, {
          type: "insertTableRow",
          blockId: `r${deletedRow}c0`,
          position: "after",
          cellTexts: ["New cell 1", "New cell 2"],
        });

        const expected: Row[] = [];
        for (let sourceRow = 0; sourceRow < rowCount; sourceRow += 1) {
          if (sourceRow === deletedRow) {
            expected.push(
              row("newc0", "New cell 1", { rowIndex: 0, column: 0 }),
              row("newc1", "New cell 2", { rowIndex: 0, column: 1 }),
            );
            continue;
          }
          expected.push(
            row(`r${sourceRow}c0`, `Row ${sourceRow} cell 1`, {
              rowIndex: sourceRow - Number(sourceRow > deletedRow),
              column: 0,
            }),
            row(`r${sourceRow}c1`, `Row ${sourceRow} cell 2`, {
              rowIndex: sourceRow - Number(sourceRow > deletedRow),
              column: 1,
            }),
          );
        }
        assert.deepEqual(
          compareWithModel(model, expected),
          [],
          `row count ${rowCount}, deleted row ${deletedRow}`,
        );
      }
    }
  });

  test("keeps row payloads on source columns across insertions and deletions", () => {
    for (const edit of ["insert", "delete"] as const) {
      for (let width = 2; width <= 4; width += 1) {
        for (let columnMask = 1; columnMask < 2 ** width; columnMask += 1) {
          if (edit === "delete" && columnMask === 2 ** width - 1) continue;
          const cells = Array.from({ length: width * 2 }, (_, index) => {
            const rowIndex = Math.floor(index / width);
            const column = index % width;
            return row(`r${rowIndex}c${column}`, `${rowIndex}:${column}`, {
              rowIndex,
              column,
            });
          });
          const model = modelOf(cells);
          for (let sourceColumn = 0; sourceColumn < width; sourceColumn += 1) {
            if ((columnMask & (1 << sourceColumn)) === 0) continue;
            expectOperation(model, {
              type: edit === "insert" ? "insertTableColumn" : "deleteTableColumn",
              blockId: `r0c${sourceColumn}`,
              position: "after",
            });
          }
          const rowAnchorColumn =
            edit === "delete"
              ? Array.from({ length: width }, (_, column) => column).find(
                  (column) => (columnMask & (1 << column)) === 0,
                )
              : 0;
          assert.ok(rowAnchorColumn !== undefined);
          expectOperation(model, {
            type: "insertTableRow",
            blockId: `r1c${rowAnchorColumn}`,
            position: "after",
            cellTexts: Array.from({ length: width }, (_, column) => `New cell ${column + 1}`),
          });

          const table = model.tables.tables[0];
          assert.ok(table);
          const actualRow = table.cells
            .filter(({ row: rowIndex }) => rowIndex === 2)
            .sort((left, right) => left.column - right.column)
            .map(({ column, paragraphs }) => [column, paragraphs]);
          const expectedRow: [number, string[]][] = [];
          let finalColumn = 0;
          for (let sourceColumn = 0; sourceColumn < width; sourceColumn += 1) {
            if (edit === "delete" && (columnMask & (1 << sourceColumn)) !== 0) continue;
            expectedRow.push([finalColumn, [`New cell ${sourceColumn + 1}`]]);
            finalColumn += 1;
            if (edit === "insert" && (columnMask & (1 << sourceColumn)) !== 0) {
              expectedRow.push([finalColumn, [""]]);
              finalColumn += 1;
            }
          }
          assert.deepEqual(actualRow, expectedRow, `${edit}, width ${width}, mask ${columnMask}`);
        }
      }
    }
  });
});
