import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import { expectTableAttrs, mergeTableAttrs } from "./attrs";

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
