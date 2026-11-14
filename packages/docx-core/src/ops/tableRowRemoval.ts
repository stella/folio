/**
 * Taking a row out of a table, directly or by resolving a suggestion. A row
 * that starts a vertical merge the next row continues does not take the
 * merged cell with it: the merge starts in the next row instead, with the
 * removed cell's properties and content, so the merged cell keeps its words.
 * Every path that removes a row goes through {@link rowsWithoutRow}.
 */
import type { Table, TableCell, TableRow } from "../model/document";
import {
  TABLE_CELL_VERTICAL_MERGES,
  tableCellVerticalMerge,
} from "../model/tableCellVerticalMerge";
import { tableGrid } from "./tableGrid";
import { DOCUMENT_OP_TYPES } from "./types";

/**
 * The cells of row `rowIndex` that own a vertical merge the next row
 * continues. Removing the row hands each on to the next row, so a suggested
 * removal of the row leaves their content as it is: accepting it keeps that
 * content in the merged cell, as removing the row directly does.
 */
export const mergeOwnersContinuedBelow = (
  table: Table,
  rowIndex: number,
): ReadonlySet<TableCell> => {
  const grid = tableGrid(table, DOCUMENT_OP_TYPES.DELETE_ROW);
  if (grid.isErr()) return new Set();
  const removed = grid.value.rows[rowIndex] ?? [];
  const owners = new Set<TableCell>();
  for (const slot of grid.value.rows[rowIndex + 1] ?? []) {
    if (slot.ownerRow !== rowIndex) continue;
    const owner = removed.find((entry) => entry.index === slot.ownerIndex);
    if (owner !== undefined) owners.add(owner.cell);
  }
  return owners;
};

/**
 * The row after `rowIndex` once that row is removed, when the removed row
 * starts a vertical merge the next row continues: the merge then starts in
 * the next row, with the removed cell's properties and content; it is no
 * merge at all when no later row continues it. Content the continuation held
 * stays after the moved content. `undefined` when the removed row starts no
 * merge that continues.
 */
const withMergesHandedOn = (
  table: Table,
  rowIndex: number,
): TableRow | undefined => {
  const grid = tableGrid(table, DOCUMENT_OP_TYPES.DELETE_ROW);
  if (grid.isErr()) return undefined;
  const removed = grid.value.rows[rowIndex] ?? [];
  const following = grid.value.rows[rowIndex + 1] ?? [];
  const later = grid.value.rows[rowIndex + 2] ?? [];
  const nextRow = table.rows[rowIndex + 1];
  if (nextRow === undefined) return undefined;
  let handed = false;
  const cells = nextRow.cells.map((cell, index) => {
    const slot = following.find((entry) => entry.index === index);
    if (
      slot === undefined ||
      slot.ownerRow !== rowIndex ||
      tableCellVerticalMerge(cell.formatting?.vMerge) !==
        TABLE_CELL_VERTICAL_MERGES.CONTINUE
    )
      return cell;
    const owner = removed.find((entry) => entry.index === slot.ownerIndex);
    if (owner === undefined) return cell;
    handed = true;
    const continues = later.some(
      (entry) => entry.start === slot.start && entry.ownerRow === rowIndex,
    );
    const { vMerge: _vMerge, ...formatting } = owner.cell.formatting ?? {};
    const held = cell.content.filter(
      (block) => block.type !== "paragraph" || block.content.length > 0,
    );
    const moved = { ...owner.cell, content: [...owner.cell.content, ...held] };
    if (continues)
      return {
        ...moved,
        formatting: {
          ...formatting,
          vMerge: TABLE_CELL_VERTICAL_MERGES.RESTART,
        },
      };
    if (Object.keys(formatting).length === 0) {
      const { formatting: _formatting, ...plain } = moved;
      return plain;
    }
    return { ...moved, formatting };
  });
  return handed ? { ...nextRow, cells } : undefined;
};

/** The table's rows without row `rowIndex`, the merges it starts handed on. */
export const rowsWithoutRow = (table: Table, rowIndex: number): TableRow[] => {
  const rows = [...table.rows];
  const next = withMergesHandedOn(table, rowIndex);
  if (next === undefined) rows.splice(rowIndex, 1);
  else rows.splice(rowIndex, 2, next);
  return rows;
};

/**
 * Whether taking rows out left a vertical merge without the cell that starts
 * it: rows that covered one grid before must still cover one after, or a
 * merged cell's content went with a row that only continued elsewhere.
 */
export const removalOrphansAMerge = (
  before: Table,
  rows: readonly TableRow[],
): boolean =>
  rows.length > 0 &&
  tableGrid(before, DOCUMENT_OP_TYPES.DELETE_ROW).isOk() &&
  tableGrid(
    { ...before, rows: [...rows] },
    DOCUMENT_OP_TYPES.DELETE_ROW,
  ).isErr();
