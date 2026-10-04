/** Logical slots, including omitted cells and vertical continuation ownership. */
import { Result } from "better-result";
import {
  TABLE_CELL_VERTICAL_MERGES,
  tableCellVerticalMerge,
} from "../model/tableCellVerticalMerge";
import type { Table, TableCell } from "../model/document";
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
export type TableGrid = { width: number; rows: GridCell[][] };

const nonnegative = (value: number) => Number.isSafeInteger(value) && value >= 0;

/** Each row covers one common grid; continuations exactly match the preceding span. */
export const tableGrid = (
  table: Table,
  opType: DocumentOpType,
): Result<TableGrid, DocumentOpRefusal> => {
  const refuse = () =>
    Result.err(
      new DocumentOpRefusal({
        opType,
        reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        message: "Table rows must cover one grid with aligned, nonoverlapping spans.",
      }),
    );
  if (table.rows.length === 0) return refuse();
  const rows: GridCell[][] = [];
  let width = table.columnWidths?.length;
  const previous = new Map<number, GridCell>();
  for (const [rowIndex, row] of table.rows.entries()) {
    let position = row.formatting?.gridBefore ?? 0;
    const after = row.formatting?.gridAfter ?? 0;
    if (!nonnegative(position) || !nonnegative(after)) return refuse();
    const cells: GridCell[] = [];
    for (const [index, cell] of row.cells.entries()) {
      if (cell.structuralChange?.type === "tableCellDeletion") continue;
      const span = cell.formatting?.gridSpan ?? 1;
      if (!nonnegative(span) || span === 0 || !nonnegative(position + span)) return refuse();
      const start = position;
      const end = start + span;
      let ownerRow = rowIndex;
      let ownerIndex = index;
      if (tableCellVerticalMerge(cell.formatting?.vMerge) === TABLE_CELL_VERTICAL_MERGES.CONTINUE) {
        const above = previous.get(start);
        const aboveMerge = tableCellVerticalMerge(above?.cell.formatting?.vMerge);
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
    const covered = position + after;
    width ??= covered;
    if (covered !== width || width === 0 || cells.length === 0) return refuse();
    rows.push(cells);
    previous.clear();
    for (const cell of cells) previous.set(cell.start, cell);
  }
  if (width === undefined) return refuse();
  return Result.ok({ width, rows });
};
