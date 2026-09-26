import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import { expectTableAttrs, mergeTableAttrs } from "./attrs";

type RemoveTableRowOptions = {
  map: TableMap;
  table: PMNode;
  tableStart: number;
};

/**
 * Remove one row, keeping every vertical merge that touches it whole: a merge
 * reaching through the row from above closes one row sooner, and a cell that
 * starts in the row and spans down moves into the row below, one row shorter.
 *
 * `prosemirror-tables` exports the same operation as `removeRow`, but its walk
 * across the row skips a cell's extra columns without moving its map index
 * along with them, so after a cell wider than one column it reads the wrong
 * slots: a merge reaching through the row further right is never shortened,
 * and the table is left one row shorter than that merge.
 */
export const removeTableRow = (
  tr: Transaction,
  { map, table, tableStart }: RemoveTableRowOptions,
  row: number,
): void => {
  let rowPosition = 0;
  for (let index = 0; index < row; index++) {
    rowPosition += table.child(index).nodeSize;
  }
  const nextRow = rowPosition + table.child(row).nodeSize;
  const mapFrom = tr.mapping.maps.length;
  tr.delete(tableStart + rowPosition, tableStart + nextRow);
  for (let column = 0; column < map.width;) {
    const index = row * map.width + column;
    const position = map.map[index];
    const cell = position === undefined ? null : table.nodeAt(position);
    if (position === undefined || !cell) {
      column += 1;
      continue;
    }
    const colspan = Math.max(1, Number(cell.attrs["colspan"]) || 1);
    const rowspan = Math.max(1, Number(cell.attrs["rowspan"]) || 1);
    if (row > 0 && position === map.map[index - map.width]) {
      tr.setNodeMarkup(tr.mapping.slice(mapFrom).map(tableStart + position), null, {
        ...cell.attrs,
        rowspan: rowspan - 1,
      });
    } else if (row + 1 < map.height && position === map.map[index + map.width]) {
      const moved = cell.type.create({ ...cell.attrs, rowspan: rowspan - 1 }, cell.content);
      const below = map.positionAt(row + 1, column, table);
      tr.insert(tr.mapping.slice(mapFrom).map(tableStart + below), moved);
    }
    column += colspan;
  }
};

/**
 * Remove every row a structural edit left without a cell of its own, closing
 * the vertical merges that reached through it over one row fewer.
 *
 * Deleting the only column a row still had a cell in, or merging whole rows
 * together, leaves a row holding nothing but the merges from above. The
 * editable model can carry such a row, but nothing else agrees with it: a
 * package spells it as a row of `w:vMerge` continuations, which the reader
 * gives a cell of its own by splitting the merge apart, so the reopened table
 * differs from the one the operation reported. Accepting the same deletion
 * tracked removes the cells one at a time and the row with its last one,
 * which is this result; the two modes have to leave the same table.
 */
export const removeRowsWithoutCells = (tr: Transaction, tablePosition: number): void => {
  for (;;) {
    const table = tr.doc.nodeAt(tablePosition);
    if (!table || table.type.spec["tableRole"] !== "table" || table.childCount < 2) {
      return;
    }
    let emptyRow = -1;
    table.forEach((row, _offset, index) => {
      if (emptyRow === -1 && row.childCount === 0) {
        emptyRow = index;
      }
    });
    if (emptyRow === -1) {
      return;
    }
    removeTableRow(
      tr,
      { map: TableMap.get(table), table, tableStart: tablePosition + 1 },
      emptyRow,
    );
  }
};

type ReconcileTableGridAfterColumnRemovalOptions = {
  tr: Transaction;
  tablePosition: number;
  previousTable: PMNode;
  removedColumn: number;
};

/**
 * Keep the editable table grid aligned with a physical column removal.
 *
 * `prosemirror-tables` owns cell topology, but Folio separately retains the
 * authored `w:tblGrid` widths and source XML on the table node. A topology
 * edit that leaves those attrs untouched serializes a new row shape against
 * the old grid; reopening then has to interpret surviving cells as spans.
 */
export const reconcileTableGridAfterColumnRemoval = ({
  tr,
  tablePosition,
  previousTable,
  removedColumn,
}: ReconcileTableGridAfterColumnRemovalOptions): void => {
  const previousColumnCount = TableMap.get(previousTable).width;
  const table = tr.doc.nodeAt(tablePosition);
  if (!table || table.type.spec["tableRole"] !== "table") {
    return;
  }

  const columnCount = TableMap.get(table).width;
  if (columnCount >= previousColumnCount) {
    return;
  }

  const removedColumnCount = previousColumnCount - columnCount;
  const previousWidths = expectTableAttrs(previousTable).columnWidths;
  const columnWidths =
    previousWidths?.length === previousColumnCount
      ? previousWidths.toSpliced(removedColumn, removedColumnCount)
      : undefined;
  setTableGrid({ tr, tablePosition, table, columnWidths, columnCount });
};

type ReconcileTableGridAfterColumnInsertionOptions = {
  tr: Transaction;
  tablePosition: number;
  previousTable: PMNode;
  insertedColumn: number;
};

/**
 * The same for a physical column insertion: the new grid column takes the
 * width of the column beside it, and the table grows by it.
 *
 * Without it the table keeps the grid it had, so every row spans one column
 * more than `w:tblGrid` declares.
 */
export const reconcileTableGridAfterColumnInsertion = ({
  tr,
  tablePosition,
  previousTable,
  insertedColumn,
}: ReconcileTableGridAfterColumnInsertionOptions): void => {
  const previousColumnCount = TableMap.get(previousTable).width;
  const table = tr.doc.nodeAt(tablePosition);
  if (!table || table.type.spec["tableRole"] !== "table") {
    return;
  }

  const columnCount = TableMap.get(table).width;
  if (columnCount <= previousColumnCount) {
    return;
  }

  const insertedColumnCount = columnCount - previousColumnCount;
  const previousWidths = expectTableAttrs(previousTable).columnWidths;
  const neighbourWidth =
    previousWidths?.[Math.max(0, Math.min(insertedColumn, previousColumnCount) - 1)] ??
    previousWidths?.[0];
  const columnWidths =
    previousWidths?.length === previousColumnCount && neighbourWidth !== undefined
      ? previousWidths.toSpliced(
          insertedColumn,
          0,
          ...Array.from({ length: insertedColumnCount }, () => neighbourWidth),
        )
      : undefined;
  setTableGrid({ tr, tablePosition, table, columnWidths, columnCount });
};

type SetTableGridOptions = {
  tr: Transaction;
  tablePosition: number;
  table: PMNode;
  columnWidths: number[] | undefined;
  columnCount: number;
};

const setTableGrid = ({
  tr,
  tablePosition,
  table,
  columnWidths,
  columnCount,
}: SetTableGridOptions): void => {
  const formatting = expectTableAttrs(table)._originalFormatting;
  const originalFormatting = formatting ? { ...formatting } : undefined;
  if (originalFormatting) {
    delete originalFormatting.gridSourceXml;
  }

  tr.setNodeMarkup(
    tablePosition,
    undefined,
    mergeTableAttrs(table, {
      columnWidths: columnWidths?.length === columnCount ? columnWidths : undefined,
      _originalFormatting: originalFormatting,
    }),
  );
};
