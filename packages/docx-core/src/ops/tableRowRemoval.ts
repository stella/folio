/**
 * Taking a row out of a table, directly or by resolving a suggestion. A row
 * that starts a vertical merge the next row continues does not take the
 * merged cell with it: the merge starts in the next row instead, with the
 * removed cell's properties and content, so the merged cell keeps its words.
 * Every path that removes a row goes through {@link rowsWithoutRow}.
 */
import type { Paragraph, Table, TableCell, TableRow } from "../model/document";
import {
  TABLE_CELL_VERTICAL_MERGES,
  tableCellVerticalMerge,
} from "../model/tableCellVerticalMerge";
import { rangeBoundaryDefinitions } from "./rangeBoundaries";
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

/**
 * The row before `rowIndex` once that row is removed, when the removed row
 * holds the only continuation of a vertical merge the row before starts: the
 * merge no row continues is no merge, and the cell starting it states none.
 * `undefined` when the removed row ends no such merge.
 */
const withMergesEnded = (
  table: Table,
  rowIndex: number,
): TableRow | undefined => {
  const grid = tableGrid(table, DOCUMENT_OP_TYPES.DELETE_ROW);
  if (grid.isErr()) return undefined;
  const previousRow = table.rows[rowIndex - 1];
  if (previousRow === undefined) return undefined;
  const removed = grid.value.rows[rowIndex] ?? [];
  const following = grid.value.rows[rowIndex + 1] ?? [];
  const ended = new Set(
    removed
      .filter(
        (slot) =>
          slot.ownerRow === rowIndex - 1 &&
          !following.some(
            (entry) =>
              entry.ownerRow === slot.ownerRow &&
              entry.ownerIndex === slot.ownerIndex,
          ),
      )
      .map((slot) => slot.ownerIndex),
  );
  if (ended.size === 0) return undefined;
  const cells = previousRow.cells.map((cell, index) => {
    if (!ended.has(index)) return cell;
    // The captured properties stated the merge.
    const {
      vMerge: _vMerge,
      sourceXml: _sourceXml,
      ...formatting
    } = cell.formatting ?? {};
    if (Object.keys(formatting).length > 0) return { ...cell, formatting };
    const { formatting: _formatting, ...plain } = cell;
    return plain;
  });
  return { ...previousRow, cells };
};

const BOUNDARIES = rangeBoundaryDefinitions();
type Boundary = Extract<
  Paragraph["content"][number],
  { type: keyof typeof BOUNDARIES }
>;

/** Every range boundary in `value`, at any depth, in source order. */
const boundariesIn = (value: unknown, out: Boundary[] = []): Boundary[] => {
  if (typeof value !== "object" || value === null) return out;
  const type: unknown = Reflect.get(value, "type");
  // SAFETY: a model node whose type names a range boundary is one.
  if (typeof type === "string" && Object.hasOwn(BOUNDARIES, type))
    out.push(value as Boundary);
  else for (const field of Object.values(value)) boundariesIn(field, out);
  return out;
};

/**
 * The range boundaries in `removed`, other than those `kept` still holds,
 * whose other end stands outside it. Removed content takes its words, not
 * the end of a range reaching past it: these ends stay where it was. A range
 * wholly inside goes with it.
 */
export const rangeEndsReachingPast = (
  removed: unknown,
  kept?: unknown,
): Boundary[] => {
  const keeps = new Set(boundariesIn(kept));
  const key = (node: Boundary) => `${BOUNDARIES[node.type].family}:${node.id}`;
  const lost = boundariesIn(removed).filter((node) => !keeps.has(node));
  return lost.filter(
    (node) => lost.filter((other) => key(other) === key(node)).length === 1,
  );
};

/**
 * `rows` with the ends of ranges reaching past the removed row
 * ({@link rangeEndsReachingPast}) where the row was: at the start of the row
 * after it, else at the end of the row before. With no row left the table
 * goes, and its removal keeps them.
 */
const withRangeEndsKept = (
  removed: TableRow,
  rows: TableRow[],
  rowIndex: number,
): TableRow[] => {
  const ends = rangeEndsReachingPast(removed, rows[rowIndex]);
  const after = rowIndex < rows.length;
  const at = after ? rowIndex : rowIndex - 1;
  const row = rows[at];
  const cellAt = after ? 0 : (row?.cells.length ?? 0) - 1;
  const cell = row?.cells[cellAt];
  if (ends.length === 0 || row === undefined || cell === undefined) return rows;
  const content = [...cell.content];
  const index = after
    ? content.findIndex((block) => block.type === "paragraph")
    : content.findLastIndex((block) => block.type === "paragraph");
  const paragraph = content[index];
  if (paragraph?.type !== "paragraph") return rows;
  content[index] = {
    ...paragraph,
    content: after
      ? [...ends, ...paragraph.content]
      : [...paragraph.content, ...ends],
  };
  const cells = [...row.cells];
  cells[cellAt] = { ...cell, content };
  const out = [...rows];
  out[at] = { ...row, cells };
  return out;
};

/**
 * The table's rows without row `rowIndex`: the merges it starts handed on,
 * those it alone continues ended, and the ends of ranges reaching past it
 * kept.
 */
export const rowsWithoutRow = (table: Table, rowIndex: number): TableRow[] => {
  const rows = [...table.rows];
  const ended = withMergesEnded(table, rowIndex);
  if (ended !== undefined) rows[rowIndex - 1] = ended;
  const next = withMergesHandedOn(table, rowIndex);
  if (next === undefined) rows.splice(rowIndex, 1);
  else rows.splice(rowIndex, 2, next);
  const removed = table.rows[rowIndex];
  return removed === undefined
    ? rows
    : withRangeEndsKept(removed, rows, rowIndex);
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
