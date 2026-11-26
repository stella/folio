import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import {
  removeTableRow,
  ROW_CLEANUP_PENDING,
  rowHoldsNoCells,
  type RowsEmptiedInBatch,
} from "../../prosemirror/tableGridMutation";

/**
 * Remove the row at `rowPos` as rejecting its insertion does: through
 * `removeTableRow`, so the vertical merges crossing it stay whole, or, for a
 * table's only row, the table with it, leaving an empty paragraph where its
 * container must keep a block.
 */
export const deleteTableRowAt = (tr: Transaction, rowPos: number): void => {
  const resolved = tr.doc.resolve(rowPos);
  const table = resolved.parent;
  if (table.type.spec["tableRole"] !== "table") {
    return;
  }
  const rowIndex = resolved.index();
  if (table.childCount > 1) {
    removeTableRow(
      tr,
      { map: TableMap.get(table), table, tableStart: resolved.start() },
      rowIndex,
    );
    return;
  }

  const tablePosition = resolved.start() - 1;
  const tableEnd = tablePosition + table.nodeSize;
  const outerResolved = tr.doc.resolve(tablePosition);
  const parent = outerResolved.parent;
  const tableIndex = outerResolved.index();
  if (parent.canReplace(tableIndex, tableIndex + 1)) {
    tr.delete(tablePosition, tableEnd);
    return;
  }
  const emptyParagraph = tr.doc.type.schema.nodes["paragraph"]?.createAndFill();
  if (
    emptyParagraph &&
    parent.canReplaceWith(
      tableIndex,
      tableIndex + 1,
      emptyParagraph.type,
      emptyParagraph.marks,
    )
  ) {
    tr.replaceWith(tablePosition, tableEnd, emptyParagraph);
  }
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
