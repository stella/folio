import type { Node as PMNode } from "prosemirror-model";

import { expectTableCellAttrs, expectTableRowAttrs } from "../../prosemirror/attrs";

/** Resolve absolute cell preferences for an unbounded fixed-layout grid (§17.4.53). */
export const fixedTableColumnWidths = (
  table: PMNode,
  grid: readonly number[],
): readonly number[] => {
  const preferred = new Map<number, number>();
  for (let rowIndex = 0; rowIndex < table.childCount; rowIndex++) {
    const row = table.child(rowIndex);
    const rowAttrs = expectTableRowAttrs(row);
    // A structural revision may supply grid columns absent from a reviewed view.
    if (rowAttrs.trIns || rowAttrs.trDel || rowAttrs.hidden) {
      return grid;
    }
    let column = Math.max(0, rowAttrs._originalFormatting?.gridBefore ?? 0);
    for (let cellIndex = 0; cellIndex < row.childCount; cellIndex++) {
      const attrs = expectTableCellAttrs(row.child(cellIndex));
      if (attrs._omittedGridSlot) {
        continue;
      }
      const width = attrs._authoredWidth;
      if (
        attrs.colspan !== 1 ||
        attrs.rowspan !== 1 ||
        attrs._preserveVMergeRestart ||
        attrs._originalFormatting?.vMerge !== undefined ||
        !width ||
        width.type !== "dxa" ||
        !Number.isFinite(width.value) ||
        width.value <= 0
      ) {
        return grid;
      }
      if (column < grid.length) {
        preferred.set(column, Math.max(preferred.get(column) ?? 0, width.value));
      }
      column++;
    }
  }
  return grid.map((width, column) => preferred.get(column) ?? width);
};
