/**
 * What a clipped table row shows. A row that continues across flow regions is
 * painted once per region and clipped to that region's slice of it, so anything
 * wholly outside the slice can stay unbuilt. Both painters share these rules.
 */

import type { TableBlock, TableRow } from "../types";
import { isFloatingImageRun, isFloatingTextBoxBlock } from "../types";

/** A vertical span in the coordinates of whatever is being painted. */
export type VisibleBand = {
  readonly top: number;
  readonly bottom: number;
};

/** The same band in the coordinates of a child that starts at `origin`. */
export const bandFromOrigin = (
  band: VisibleBand,
  origin: number,
): VisibleBand => ({
  top: band.top - origin,
  bottom: band.bottom - origin,
});

export const isOutsideBand = (
  band: VisibleBand | undefined,
  top: number,
  bottom: number,
): boolean => band !== undefined && (bottom <= band.top || top >= band.bottom);

/**
 * Where a row's paint ends: a merged cell belongs to the row that opens it and
 * reaches through every row it spans. `rowTops` has `rows.length + 1` entries.
 */
export const tableRowPaintBottom = (
  row: TableRow,
  rowIndex: number,
  rowTops: readonly number[],
): number => {
  let widestSpan = 1;
  for (const cell of row.cells) {
    widestSpan = Math.max(widestSpan, cell.rowSpan ?? 1);
  }
  const lastTop = rowTops.length - 1;
  return rowTops[Math.min(lastTop, rowIndex + widestSpan)] ?? 0;
};

/** Floating cell content can paint outside the row it is anchored in. */
export function tableHasFloatingCellContent(block: TableBlock): boolean {
  for (const row of block.rows) {
    for (const cell of row.cells) {
      for (const cellBlock of cell.blocks) {
        if (cellBlock.kind === "textBox" && isFloatingTextBoxBlock(cellBlock)) {
          return true;
        }
        if (
          cellBlock.kind === "paragraph" &&
          cellBlock.runs.some(
            (run) => run.kind === "image" && isFloatingImageRun(run),
          )
        ) {
          return true;
        }
        if (
          cellBlock.kind === "table" &&
          tableHasFloatingCellContent(cellBlock)
        ) {
          return true;
        }
      }
    }
  }
  return false;
}
