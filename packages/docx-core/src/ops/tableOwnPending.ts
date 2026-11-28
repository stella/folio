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

type OwnPendingTableInsertion =
  | { type: "wholeTable" }
  | { type: "cells"; revisionIds: number[]; blockId: string };

const insertedBy = (cell: TableCell, owner: RevisionOwner): boolean =>
  cell.structuralChange?.type === "tableCellInsertion" &&
  isOwnRecord(cell.structuralChange.info, owner);

/**
 * What a tracked deletion of a table removes outright: the whole table when
 * every row is the same author's pending insertion, or cells that author's
 * pending column insertion added, with no one else's suggestion in them.
 * Deleting over one's own insertion withdraws it instead of recording a
 * deletion of the addition. A partial result supplies the ids to reject and a
 * paragraph of a staying cell that addresses the table afterwards.
 */
export const ownPendingTableInsertion = (
  document: Document,
  op: DeleteTableOp,
): OwnPendingTableInsertion | undefined => {
  const owner = op.revision;
  if (owner === undefined) return undefined;
  const located = locateTableRow(document, op);
  if (located.isErr()) return undefined;
  const table: Table = located.value.table;
  if (op.expected !== undefined && !equalForStaleness(table, op.expected))
    return undefined;
  if (
    table.rows.length > 0 &&
    table.rows.every(
      (row) =>
        row.structuralChange?.type === "tableRowInsertion" &&
        isOwnRecord(row.structuralChange.info, owner),
    ) &&
    recordedOnlyBy(table, owner)
  )
    return { type: "wholeTable" };
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
    type: "cells",
    revisionIds: [
      ...(table.propertyChanges ?? []).map(({ info }) => info.id),
      ...inserted.flatMap(({ structuralChange }) =>
        structuralChange === undefined ? [] : [structuralChange.info.id],
      ),
    ],
    blockId: paragraph.paraId,
  };
};
