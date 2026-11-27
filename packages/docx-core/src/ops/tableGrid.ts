/** Logical slots, including omitted cells and vertical continuation ownership. */
import { Result } from "better-result";

import type { Table, TableCell } from "../model/document";
import { maxOf } from "../model/extremes";
import {
  TABLE_CELL_VERTICAL_MERGES,
  tableCellVerticalMerge,
} from "../model/tableCellVerticalMerge";
import { MAX_TABLE_GRID_COLUMNS } from "../model/tableGridLimits";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { DocumentOpType } from "./types";

export type GridCell = {
  cell: TableCell;
  row: number;
  index: number;
  start: number;
  end: number;
  ownerRow: number;
  ownerIndex: number;
};
export type TableGrid = {
  width: number;
  rows: GridCell[][];
  /**
   * Whether every row covers the whole grid and the table's column widths
   * name every grid column. A row that ends early is read with the empty
   * columns it implies after it, as a trailing `gridAfter`.
   */
  stated: boolean;
};

const nonnegative = (value: number) =>
  Number.isSafeInteger(value) && value >= 0;

/**
 * Each row covers one common grid, as wide as the widest row or the table's
 * column widths; continuations exactly match the preceding span.
 */
export const tableGrid = (
  table: Table,
  opType: DocumentOpType,
): Result<TableGrid, DocumentOpRefusal> => {
  const refuse = () =>
    Result.err(
      new DocumentOpRefusal({
        opType,
        reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        message:
          "Table rows must cover one grid with aligned, nonoverlapping spans.",
      }),
    );
  if (table.rows.length === 0) return refuse();
  const rows: GridCell[][] = [];
  const covered: number[] = [];
  const previous = new Map<number, GridCell>();
  for (const [rowIndex, row] of table.rows.entries()) {
    let position = row.formatting?.gridBefore ?? 0;
    const after = row.formatting?.gridAfter ?? 0;
    if (!nonnegative(position) || !nonnegative(after)) return refuse();
    const cells: GridCell[] = [];
    for (const [index, cell] of row.cells.entries()) {
      // A pending cell deletion still holds its columns: the grid narrows only
      // when the deletion is accepted.
      const span = cell.formatting?.gridSpan ?? 1;
      if (!nonnegative(span) || span === 0 || !nonnegative(position + span))
        return refuse();
      const start = position;
      const end = start + span;
      let ownerRow = rowIndex;
      let ownerIndex = index;
      if (
        tableCellVerticalMerge(cell.formatting?.vMerge) ===
        TABLE_CELL_VERTICAL_MERGES.CONTINUE
      ) {
        const above = previous.get(start);
        const aboveMerge = tableCellVerticalMerge(
          above?.cell.formatting?.vMerge,
        );
        if (
          !above ||
          above.end !== end ||
          (aboveMerge !== TABLE_CELL_VERTICAL_MERGES.RESTART &&
            aboveMerge !== TABLE_CELL_VERTICAL_MERGES.CONTINUE)
        )
          return refuse();
        ownerRow = above.ownerRow;
        ownerIndex = above.ownerIndex;
      }
      cells.push({
        cell,
        row: rowIndex,
        index,
        start: position,
        end: position + span,
        ownerRow,
        ownerIndex,
      });
      position += span;
    }
    // A row may hold no cells (`CT_Row` takes zero or more): it covers what
    // it states it omits.
    covered.push(position + after);
    rows.push(cells);
    previous.clear();
    for (const cell of cells) previous.set(cell.start, cell);
  }
  const width = maxOf(covered, table.columnWidths?.length ?? 0);
  if (width === 0 || width > MAX_TABLE_GRID_COLUMNS) return refuse();
  const stated =
    covered.every((columns) => columns === width) &&
    (table.columnWidths === undefined || table.columnWidths.length === width);
  return Result.ok({ width, rows, stated });
};

/** Whether a row starts a cell of its own rather than only continuing merges. */
const startsCell = (cells: readonly TableCell[]) =>
  cells.some(
    (cell) =>
      tableCellVerticalMerge(cell.formatting?.vMerge) !==
      TABLE_CELL_VERTICAL_MERGES.CONTINUE,
  );

type EditRemovesRowOptions = {
  /** The row's cells before the edit. */
  before: readonly TableCell[];
  /** What the edit leaves of them. */
  after: readonly TableCell[];
};

/**
 * Whether an edit that leaves a row with `after` removes the row: it took
 * every cell the row had, or every cell the row started, which would leave
 * a row of nothing but merges from above. A reader cannot hold such a row:
 * it splits the merge apart to give the row a cell. The merges above close
 * over one row fewer.
 */
export const editRemovesRow = ({ before, after }: EditRemovesRowOptions) =>
  (before.length > 0 && after.length === 0) ||
  (startsCell(before) && !startsCell(after));
