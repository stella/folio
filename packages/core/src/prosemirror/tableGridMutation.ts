import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { removeColumn, TableMap } from "prosemirror-tables";

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
 * Whether a batch of structural edits may have left rows without cells, to
 * close once the whole batch has run.
 */
export type RowsEmptiedInBatch = { pending: boolean };

/**
 * Close the rows a batch left without cells, once the whole batch has run: a
 * later operation of the same batch can still give such a row a cell (a column
 * inserted beside a merge), as accepting the batch tracked does.
 *
 * Every table is visited rather than the ones the deletions named: a later
 * operation of the batch can move a table whole, as a merge carries a nested
 * table into the merged cell, and a position mapped through that move no
 * longer finds it. Tables go from the last to the first, so closing rows in
 * one never moves a table still to visit.
 */
export const removeRowsWithoutCellsAfterBatch = (
  tr: Transaction,
  emptied: RowsEmptiedInBatch,
): void => {
  if (!emptied.pending) {
    return;
  }
  const tables: number[] = [];
  tr.doc.descendants((node, position) => {
    if (node.type.spec["tableRole"] === "table") {
      tables.push(position);
    }
    return !node.isTextblock;
  });
  for (const position of tables.toReversed()) {
    removeRowsWithoutCells(tr, position);
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
  const neighbourWidth = insertedColumnWidth(previousTable, insertedColumn);
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

/** The width a column inserted at `insertedColumn` takes: its left neighbour's, when the grid is known. */
const insertedColumnWidth = (table: PMNode, insertedColumn: number): number | undefined => {
  const columnCount = TableMap.get(table).width;
  const widths = expectTableAttrs(table).columnWidths;
  if (widths?.length !== columnCount) {
    return undefined;
  }
  return widths[Math.max(0, Math.min(insertedColumn, columnCount) - 1)] ?? widths[0];
};

/**
 * Give a cell whose span grew or shrank the matching rendered width: its
 * resolved twips width moves by `delta`. Layout falls back to cell widths
 * when the grid is unknown, so a merged cell must not keep its old one. A
 * percentage width, and the authored `w:tcW` kept beside it, stay as they are.
 */
const resizeCellWidth = (tr: Transaction, position: number, delta: number): void => {
  const cell = tr.doc.nodeAt(position);
  const width: unknown = cell?.attrs["width"];
  const widthType: unknown = cell?.attrs["widthType"];
  if (
    !cell ||
    delta === 0 ||
    typeof width !== "number" ||
    (widthType !== null && widthType !== undefined && widthType !== "dxa")
  ) {
    return;
  }
  tr.setNodeMarkup(position, null, { ...cell.attrs, width: Math.max(1, width + delta) });
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

type InsertTableColumnOptions = {
  tr: Transaction;
  tablePosition: number;
  /** The grid column the new column takes; the columns from it on move right. */
  column: number;
  /** The grid column whose cells the new cells copy their formatting from. */
  templateColumn: number;
  /** A new, empty cell shaped after `template`, the row's cell in the template column. */
  createCell: (template: PMNode | null) => PMNode;
};

/**
 * Insert a grid column, walking the table map: each row gains a cell at the
 * column's place in the grid (after any cell a merge from a row above holds
 * there), a cell merged across that place widens instead, and `w:tblGrid`
 * gains the column.
 */
export const insertTableColumn = ({
  tr,
  tablePosition,
  column,
  templateColumn,
  createCell,
}: InsertTableColumnOptions): void => {
  const table = tr.doc.nodeAt(tablePosition);
  if (!table || table.type.spec["tableRole"] !== "table") {
    return;
  }
  const map = TableMap.get(table);
  const tableStart = tablePosition + 1;
  const mapFrom = tr.mapping.maps.length;
  const addedWidth = insertedColumnWidth(table, column) ?? 0;
  for (let row = 0; row < map.height; row++) {
    const index = row * map.width + column;
    const position = map.map[index];
    const straddling =
      column > 0 && column < map.width && position !== undefined && map.map[index - 1] === position
        ? table.nodeAt(position)
        : null;
    if (straddling && position !== undefined) {
      const offset = column - map.colCount(position);
      const colwidth: unknown = straddling.attrs["colwidth"];
      const mapped = tr.mapping.slice(mapFrom).map(tableStart + position);
      tr.setNodeMarkup(mapped, null, {
        ...straddling.attrs,
        colspan: (Number(straddling.attrs["colspan"]) || 1) + 1,
        colwidth: Array.isArray(colwidth) ? colwidth.toSpliced(offset, 0, 0) : colwidth,
      });
      resizeCellWidth(tr, mapped, addedWidth);
      row += (Number(straddling.attrs["rowspan"]) || 1) - 1;
      continue;
    }
    const templatePosition = map.map[row * map.width + templateColumn];
    const template = templatePosition === undefined ? null : table.nodeAt(templatePosition);
    tr.insert(
      tr.mapping.slice(mapFrom).map(tableStart + map.positionAt(row, column, table)),
      createCell(template),
    );
  }
  reconcileTableGridAfterColumnInsertion({
    tr,
    tablePosition,
    previousTable: table,
    insertedColumn: column,
  });
};

/**
 * Remove grid columns [`left`, `right`), walking the table map: a cell only
 * in them goes, a cell merged across them narrows, `w:tblGrid` loses them,
 * and a row left without a cell of its own goes too.
 */
export const removeTableColumns = (
  tr: Transaction,
  tablePosition: number,
  left: number,
  right: number,
): void => {
  for (let column = right - 1; column >= left; column--) {
    const table = tr.doc.nodeAt(tablePosition);
    if (!table || table.type.spec["tableRole"] !== "table") {
      return;
    }
    const map = TableMap.get(table);
    const widths = expectTableAttrs(table).columnWidths;
    const removedWidth = widths?.length === map.width ? (widths[column] ?? 0) : 0;
    // The cells merged across the column narrow rather than go.
    const narrowed = new Set<number>();
    for (let row = 0; row < map.height; row++) {
      const index = row * map.width + column;
      const position = map.map[index];
      if (
        position !== undefined &&
        ((column > 0 && map.map[index - 1] === position) ||
          (column < map.width - 1 && map.map[index + 1] === position))
      ) {
        narrowed.add(position);
      }
    }
    const mapFrom = tr.mapping.maps.length;
    removeColumn(
      tr,
      {
        map,
        table,
        tableStart: tablePosition + 1,
        left: column,
        right: column + 1,
        top: 0,
        bottom: map.height,
      },
      column,
    );
    for (const position of narrowed) {
      resizeCellWidth(
        tr,
        tr.mapping.slice(mapFrom).map(tablePosition + 1 + position),
        -removedWidth,
      );
    }
    reconcileTableGridAfterColumnRemoval({
      tr,
      tablePosition,
      previousTable: table,
      removedColumn: column,
    });
  }
  removeRowsWithoutCells(tr, tablePosition);
};
