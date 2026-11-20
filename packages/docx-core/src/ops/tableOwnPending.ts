import type { Document, Table, TableCell } from "../model/document";
import { storyParagraphs } from "./blocks";
import { equalForStaleness } from "./equality";
import { locateTableRow } from "./tableLocation";
import type { DeleteTableOp } from "./types";

/** Every tracked record in `value`, at any depth, is `author`'s. */
export const recordedOnlyBy = (value: unknown, author: string): boolean => {
  if (Array.isArray(value))
    return value.every((item) => recordedOnlyBy(item, author));
  if (typeof value !== "object" || value === null) return true;
  const info: unknown = Reflect.get(value, "info");
  if (
    typeof info === "object" &&
    info !== null &&
    Reflect.get(info, "author") !== undefined &&
    Reflect.get(info, "author") !== author
  )
    return false;
  return Object.values(value).every((item) => recordedOnlyBy(item, author));
};

const insertedBy = (cell: TableCell, author: string): boolean =>
  cell.structuralChange?.type === "tableCellInsertion" &&
  cell.structuralChange.info.author === author;

/**
 * What a tracked deletion of a table removes outright first: cells the same
 * author's pending insertion added, with no one else's suggestion in them,
 * and that author's pending changes of the table's properties (a column
 * insertion records both), as deleting over one's own insertion removes it.
 * The ids to reject, and a paragraph of a cell that stays to address the
 * table by afterwards; `undefined` when the table holds no such insertion or
 * any other pending structure.
 */
export const ownPendingTableInsertion = (
  document: Document,
  op: DeleteTableOp,
): { revisionIds: number[]; blockId: string } | undefined => {
  const author = op.revision?.author;
  if (author === undefined) return undefined;
  const located = locateTableRow(document, op);
  if (located.isErr()) return undefined;
  const table: Table = located.value.table;
  if (op.expected !== undefined && !equalForStaleness(table, op.expected))
    return undefined;
  const cells = table.rows.flatMap((row) => row.cells);
  const inserted = cells.filter((cell) => insertedBy(cell, author));
  if (
    inserted.length === 0 ||
    table.rows.some((row) => row.structuralChange !== undefined) ||
    cells.some(
      (cell) =>
        (cell.structuralChange !== undefined && !insertedBy(cell, author)) ||
        (insertedBy(cell, author) && !recordedOnlyBy(cell, author)),
    ) ||
    (table.propertyChanges ?? []).some(({ info }) => info.author !== author)
  )
    return undefined;
  const staying = cells.find((cell) => !insertedBy(cell, author));
  const paragraph =
    staying === undefined
      ? undefined
      : storyParagraphs({ content: staying.content }).at(0)?.paragraph;
  if (paragraph?.paraId === undefined) return undefined;
  return {
    revisionIds: [
      ...(table.propertyChanges ?? []).map(({ info }) => info.id),
      ...inserted.flatMap(({ structuralChange }) =>
        structuralChange === undefined ? [] : [structuralChange.info.id],
      ),
    ],
    blockId: paragraph.paraId,
  };
};
