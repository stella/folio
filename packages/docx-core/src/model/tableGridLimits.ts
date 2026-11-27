import type { TableRow } from "./content";

/** The supported width when materializing table grids and their history. */
export const MAX_TABLE_GRID_COLUMNS = 4096;

/** The widest row, including the slots it omits on either edge. */
export const gridColumnsInRows = (rows: readonly TableRow[]): number => {
  let widest = 0;
  for (const row of rows) {
    let columns =
      (row.formatting?.gridBefore ?? 0) + (row.formatting?.gridAfter ?? 0);
    for (const cell of row.cells) columns += cell.formatting?.gridSpan ?? 1;
    widest = Math.max(widest, columns);
  }
  return widest;
};
