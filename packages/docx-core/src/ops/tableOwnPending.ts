import type { Document, Table, TableCell } from "../model/document";
import { storyParagraphs } from "./blocks";
import { equalForStaleness } from "./equality";
import {
  isOwnRecord,
  recordedOnlyBy,
  type RevisionOwner,
} from "./revisionOwner";
import { locateTableRow } from "./tableLocation";
import type { DeleteTableOp } from "./types";

const insertedBy = (cell: TableCell, owner: RevisionOwner): boolean =>
  cell.structuralChange?.type === "tableCellInsertion" &&
  isOwnRecord(cell.structuralChange.info, owner);

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
  const owner = op.revision;
  if (owner === undefined) return undefined;
  const located = locateTableRow(document, op);
  if (located.isErr()) return undefined;
  const table: Table = located.value.table;
  if (op.expected !== undefined && !equalForStaleness(table, op.expected))
    return undefined;
  const cells = table.rows.flatMap((row) => row.cells);
  const inserted = cells.filter((cell) => insertedBy(cell, owner));
  if (
    inserted.length === 0 ||
    table.rows.some((row) => row.structuralChange !== undefined) ||
    cells.some(
      (cell) =>
        (cell.structuralChange !== undefined && !insertedBy(cell, owner)) ||
        (insertedBy(cell, owner) && !recordedOnlyBy(cell, owner)),
    ) ||
    (table.propertyChanges ?? []).some(({ info }) => !isOwnRecord(info, owner))
  )
    return undefined;
  const staying = cells.find((cell) => !insertedBy(cell, owner));
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
