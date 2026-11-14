import { panic } from "better-result";
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
      tr.setNodeMarkup(
        tr.mapping.slice(mapFrom).map(tableStart + position),
        null,
        {
          ...cell.attrs,
          rowspan: rowspan - 1,
        },
      );
    } else if (
      row + 1 < map.height &&
      position === map.map[index + map.width]
    ) {
      const moved = cell.type.create(
        { ...cell.attrs, rowspan: rowspan - 1 },
        cell.content,
      );
      const below = map.positionAt(row + 1, column, table);
      tr.insert(tr.mapping.slice(mapFrom).map(tableStart + below), moved);
    }
    column += colspan;
  }
};

/** A placeholder for an absent grid position is not a cell a row can keep. */
export const rowHoldsNoCells = (row: PMNode, besides?: PMNode): boolean => {
  let holdsCell = false;
  row.forEach((cell) => {
    holdsCell ||= cell !== besides && cell.attrs["_omittedGridSlot"] == null;
  });
  return !holdsCell;
};

const omittedSlotOf = (cell: PMNode | null | undefined) =>
  cell?.attrs["_omittedGridSlot"] as "before" | "after" | null | undefined;

/**
 * The placeholder a new column's place in a row falls to, if any. Nothing can
 * stand after what a row omits at its end or before what it omits at its start:
 * a cell there would move on save, so the omission grows instead.
 */
export const omittedSlotTakingColumn = (
  map: TableMap,
  table: PMNode,
  row: number,
  column: number,
): number | null => {
  const left = column > 0 ? map.map[row * map.width + column - 1] : undefined;
  if (left !== undefined && omittedSlotOf(table.nodeAt(left)) === "after") {
    return left;
  }
  const right =
    column < map.width ? map.map[row * map.width + column] : undefined;
  return right !== undefined && omittedSlotOf(table.nodeAt(right)) === "before"
    ? right
    : null;
};

/**
 * Size what rows omit to the width their cells now reach. A column whose cells
 * went takes the same column out of a row that omits it; removing a tracked
 * column's last cell does not otherwise reach such a row.
 */
export const reconcileOmittedGridSlots = (
  tr: Transaction,
  tablePosition: number,
  previousTable: PMNode,
  removedColumn: number,
): void => {
  const table = tr.doc.nodeAt(tablePosition);
  if (!table || table.type.spec["tableRole"] !== "table") return;
  const reached = Array.from({ length: table.childCount }, () => 0);
  let omits = false;
  table.forEach((row, _offset, rowIndex) => {
    row.forEach((cell) => {
      if (omittedSlotOf(cell)) {
        omits = true;
        return;
      }
      const colspan = Number(cell.attrs["colspan"]) || 1;
      const rowspan = Number(cell.attrs["rowspan"]) || 1;
      for (let at = rowIndex; at < rowIndex + rowspan; at++) {
        if (at < reached.length) reached[at] = (reached[at] ?? 0) + colspan;
      }
    });
  });
  if (!omits) return;
  const width = Math.max(...reached);
  const previousMap = TableMap.get(previousTable);
  const rowPositions: number[] = [];
  table.forEach((_row, offset) =>
    rowPositions.push(tablePosition + 1 + offset),
  );
  for (let rowIndex = table.childCount - 1; rowIndex >= 0; rowIndex--) {
    const row = table.child(rowIndex);
    let excess = -(width - (reached[rowIndex] ?? 0));
    row.forEach((cell) => {
      if (omittedSlotOf(cell)) {
        excess += Number(cell.attrs["colspan"]) || 1;
      }
    });
    if (excess <= 0) continue;
    // First the omission at the column's place, then the latest.
    const atPlace =
      previousMap.map[rowIndex * previousMap.width + removedColumn];
    const placeKind =
      atPlace === undefined
        ? undefined
        : omittedSlotOf(previousTable.nodeAt(atPlace));
    const slots: { node: PMNode; position: number }[] = [];
    row.forEach((cell, offset) => {
      if (omittedSlotOf(cell)) {
        slots.push({
          node: cell,
          position: (rowPositions[rowIndex] ?? 0) + 1 + offset,
        });
      }
    });
    const order = slots.toSorted(
      (left, right) =>
        Number(omittedSlotOf(right.node) === placeKind) -
          Number(omittedSlotOf(left.node) === placeKind) ||
        right.position - left.position,
    );
    // Apply from the row's end, so an earlier position still holds.
    const changes = new Map<number, number>();
    for (const { node, position } of order) {
      if (excess <= 0) break;
      const colspan = Number(node.attrs["colspan"]) || 1;
      const taken = Math.min(colspan, excess);
      changes.set(position, colspan - taken);
      excess -= taken;
    }
    for (const { node, position } of slots.toReversed()) {
      const next = changes.get(position);
      if (next === undefined) continue;
      if (next === 0) tr.delete(position, position + node.nodeSize);
      else tr.setNodeAttribute(position, "colspan", next);
    }
  }
  removeRowsWithoutCells(tr, tablePosition, previousTable);
};

/** Whether a batch of structural deletions may have emptied rows. */
export type RowsEmptiedInBatch = { pending: boolean };

const ROW_CLEANUP_PENDING = "pending";

type MarkRowEmptiedInBatchOptions = {
  tr: Transaction;
  rowPosition: number;
  emptied: RowsEmptiedInBatch;
};

/** A transient row attribute follows node moves; conversion never serializes it. */
export const markRowEmptiedInBatch = ({
  tr,
  rowPosition,
  emptied,
}: MarkRowEmptiedInBatchOptions): void => {
  const row = tr.doc.nodeAt(rowPosition);
  if (!row || row.type.spec.attrs?.["_batchRowCleanup"] === undefined) {
    panic("Batch row cleanup requires the table row cleanup attribute");
  }
  tr.setNodeAttribute(rowPosition, "_batchRowCleanup", ROW_CLEANUP_PENDING);
  emptied.pending = true;
};

type MarkRowsEmptiedInBatchOptions = {
  tr: Transaction;
  tablePosition: number;
  previousTable: PMNode;
  emptied: RowsEmptiedInBatch;
};

/** Column removal keeps row indices stable: mark only newly cell-less rows. */
export const markRowsEmptiedInBatch = ({
  tr,
  tablePosition,
  previousTable,
  emptied,
}: MarkRowsEmptiedInBatchOptions): void => {
  const table = tr.doc.nodeAt(tablePosition);
  if (!table || table.type.spec["tableRole"] !== "table") return;
  table.forEach((row, offset, index) => {
    if (rowHoldsNoCells(row) && !rowHoldsNoCells(previousTable.child(index))) {
      markRowEmptiedInBatch({
        tr,
        rowPosition: tablePosition + 1 + offset,
        emptied,
      });
    }
  });
};

type RemoveEmptyRowsOptions = {
  tr: Transaction;
  tablePosition: number;
  shouldRemove: (row: PMNode) => boolean;
};

const removeEmptyRows = ({
  tr,
  tablePosition,
  shouldRemove,
}: RemoveEmptyRowsOptions): void => {
  for (;;) {
    const table = tr.doc.nodeAt(tablePosition);
    if (
      !table ||
      table.type.spec["tableRole"] !== "table" ||
      table.childCount < 2
    )
      return;
    let emptyRow = -1;
    table.forEach((row, _offset, index) => {
      if (emptyRow === -1 && rowHoldsNoCells(row) && shouldRemove(row))
        emptyRow = index;
    });
    if (emptyRow === -1) return;
    removeTableRow(
      tr,
      { map: TableMap.get(table), table, tableStart: tablePosition + 1 },
      emptyRow,
    );
  }
};

/** Remove only rows emptied by deletions in this transaction, then clear remaining markers. */
const removeBatchRows = (tr: Transaction, tablePosition: number): void => {
  removeEmptyRows({
    tr,
    tablePosition,
    shouldRemove: (row) =>
      row.attrs["_batchRowCleanup"] === ROW_CLEANUP_PENDING,
  });
  tr.doc.nodeAt(tablePosition)?.forEach((row, offset) => {
    if (row.attrs["_batchRowCleanup"] === ROW_CLEANUP_PENDING) {
      tr.setNodeAttribute(tablePosition + 1 + offset, "_batchRowCleanup", null);
    }
  });
};

/** Sweep current tables in reverse so moves cannot invalidate cleanup targets. */
export const removeRowsWithoutCellsAfterBatch = (
  tr: Transaction,
  emptied: RowsEmptiedInBatch,
): void => {
  if (!emptied.pending) return;
  const tables: number[] = [];
  tr.doc.descendants((node, position) => {
    if (node.type.spec["tableRole"] === "table") tables.push(position);
    return !node.isTextblock;
  });
  for (const position of tables.toReversed()) removeBatchRows(tr, position);
};

type RemoveRowsWithoutCellsInRangeOptions = {
  tr: Transaction;
  from: number;
  to: number;
};

/** Close marked rows before their owning cell is captured into a continuation payload. */
export const removeRowsWithoutCellsInRange = ({
  tr,
  from,
  to,
}: RemoveRowsWithoutCellsInRangeOptions): void => {
  const tables: number[] = [];
  tr.doc.nodesBetween(from, to, (node, position) => {
    if (position >= from && node.type.spec["tableRole"] === "table")
      tables.push(position);
    return !node.isTextblock;
  });
  for (const position of tables.toReversed()) removeBatchRows(tr, position);
};

/**
 * Close the rows an immediate structural edit left without cells. A row written
 * with none is no row this edit emptied, and it stays.
 */
export const removeRowsWithoutCells = (
  tr: Transaction,
  tablePosition: number,
  previousTable: PMNode,
): void => {
  const table = tr.doc.nodeAt(tablePosition);
  if (!table || table.type.spec["tableRole"] !== "table") return;
  const emptied: number[] = [];
  table.forEach((row, _offset, index) => {
    const before = previousTable.maybeChild(index);
    if (rowHoldsNoCells(row) && before && !rowHoldsNoCells(before)) {
      emptied.push(index);
    }
  });
  for (const rowIndex of emptied.toReversed()) {
    const current = tr.doc.nodeAt(tablePosition);
    if (!current || current.childCount < 2) return;
    removeTableRow(
      tr,
      {
        map: TableMap.get(current),
        table: current,
        tableStart: tablePosition + 1,
      },
      rowIndex,
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
  // A grid wider than the rows keeps the columns no row reaches.
  const columnWidths =
    previousWidths !== undefined && previousWidths.length >= previousColumnCount
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
    previousWidths !== undefined &&
    previousWidths.length >= previousColumnCount &&
    neighbourWidth !== undefined
      ? previousWidths.toSpliced(
          insertedColumn,
          0,
          ...Array.from({ length: insertedColumnCount }, () => neighbourWidth),
        )
      : undefined;
  setTableGrid({ tr, tablePosition, table, columnWidths, columnCount });
};

/** The width a column inserted at `insertedColumn` takes: its left neighbour's, when the grid is known. */
export const insertedColumnWidth = (
  table: PMNode,
  insertedColumn: number,
): number | undefined => {
  const columnCount = TableMap.get(table).width;
  const widths = expectTableAttrs(table).columnWidths;
  if (widths === undefined || widths.length < columnCount) {
    return undefined;
  }
  return (
    widths[Math.max(0, Math.min(insertedColumn, columnCount) - 1)] ?? widths[0]
  );
};

/**
 * Give a cell whose span grew or shrank the matching rendered width: its
 * resolved twips width moves by `delta`. Layout falls back to cell widths
 * when the grid is unknown, so a merged cell must not keep its old one. A
 * percentage width, and the authored `w:tcW` kept beside it, stay as they are.
 */
const resizeCellWidth = (
  tr: Transaction,
  position: number,
  delta: number,
): void => {
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
  tr.setNodeMarkup(position, null, {
    ...cell.attrs,
    width: Math.max(1, width + delta),
  });
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
      columnWidths:
        columnWidths !== undefined && columnWidths.length >= columnCount
          ? columnWidths
          : undefined,
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
    const omitted = omittedSlotTakingColumn(map, table, row, column);
    if (omitted !== null) {
      const slot = table.nodeAt(omitted);
      tr.setNodeAttribute(
        tr.mapping.slice(mapFrom).map(tableStart + omitted),
        "colspan",
        (Number(slot?.attrs["colspan"]) || 1) + 1,
      );
      continue;
    }
    const straddling =
      column > 0 &&
      column < map.width &&
      position !== undefined &&
      map.map[index - 1] === position
        ? table.nodeAt(position)
        : null;
    if (straddling && position !== undefined) {
      const offset = column - map.colCount(position);
      const colwidth: unknown = straddling.attrs["colwidth"];
      const mapped = tr.mapping.slice(mapFrom).map(tableStart + position);
      tr.setNodeMarkup(mapped, null, {
        ...straddling.attrs,
        colspan: (Number(straddling.attrs["colspan"]) || 1) + 1,
        colwidth: Array.isArray(colwidth)
          ? colwidth.toSpliced(offset, 0, 0)
          : colwidth,
      });
      resizeCellWidth(tr, mapped, addedWidth);
      row += (Number(straddling.attrs["rowspan"]) || 1) - 1;
      continue;
    }
    const templatePosition = map.map[row * map.width + templateColumn];
    const templateCell =
      templatePosition === undefined ? null : table.nodeAt(templatePosition);
    const template = omittedSlotOf(templateCell) ? null : templateCell;
    tr.insert(
      tr.mapping
        .slice(mapFrom)
        .map(tableStart + map.positionAt(row, column, table)),
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
  const previousTable = tr.doc.nodeAt(tablePosition);
  if (!previousTable) return;
  for (let column = right - 1; column >= left; column--) {
    const table = tr.doc.nodeAt(tablePosition);
    if (!table || table.type.spec["tableRole"] !== "table") {
      return;
    }
    const map = TableMap.get(table);
    const widths = expectTableAttrs(table).columnWidths;
    const removedWidth =
      widths?.length === map.width ? (widths[column] ?? 0) : 0;
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
  removeRowsWithoutCells(tr, tablePosition, previousTable);
};
