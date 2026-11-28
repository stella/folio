import type { TableRow } from "./content";

/** The supported width when materializing table grids and their history. */
export const MAX_TABLE_GRID_COLUMNS = 4096;

/** A zero gridSpan occupies one column (MS-OI29500 §2.1.130). */
export const gridSpanInColumns = (span: number | undefined): number =>
  span === undefined || span === 0 ? 1 : span;

/** The widest row, including the slots it omits on either edge. */
export const gridColumnsInRows = (rows: readonly TableRow[]): number => {
  let widest = 0;
  for (const row of rows) {
    let columns =
      (row.formatting?.gridBefore ?? 0) + (row.formatting?.gridAfter ?? 0);
    for (const cell of row.cells)
      columns += gridSpanInColumns(cell.formatting?.gridSpan);
    widest = Math.max(widest, columns);
  }
  return widest;
};
