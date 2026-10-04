/** Table-level review after inline changes, before paragraph-mark joins. */
import { Result } from "better-result";

import type {
  BlockContent,
  Document,
  Table,
  TableCell,
  TableCellPropertyChange,
  TableRow,
  TableStructuralChangeInfo,
} from "../model/document";
import { storyBody } from "./blocks";
import { combineEdits, type DocumentEdit } from "./edits";
import type { ApplyOps } from "./resolve";
import { locateTableRow, tableRowAnchor } from "./tableLocation";
import { tableGrid } from "./tableGrid";
import {
  DOCUMENT_OP_TYPES,
  REVISION_DECISIONS,
  type OpStory,
  type ResolveRevisionOp,
} from "./types";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";

const collectTables = (blocks: readonly BlockContent[], out: Table[]): void => {
  for (const block of blocks) {
    switch (block.type) {
      case "table":
        for (const row of block.rows) {
          for (const cell of row.cells) collectTables(cell.content, out);
        }
        out.push(block);
        break;
      case "blockSdt":
      case "blockCustomXml":
        collectTables(block.content, out);
        break;
      case "paragraph":
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        break;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  }
};
const tablesIn = (document: Document, story: OpStory): Table[] => {
  const out: Table[] = [];
  collectTables(storyBody(document, story).content, out);
  return out;
};

const propertyIds = (changes: readonly { info: { id: number } }[] | undefined): number[] =>
  changes?.map(({ info }) => info.id) ?? [];
const cellStructuralIds = (cell: TableCell): number[] =>
  cell.structuralChange === undefined ? [] : [cell.structuralChange.info.id];

/** Revision ids the table resolver can reach, including properties and cell marks. */
export const reachableRowIds = (document: Document, story: OpStory): number[] => {
  const ids: number[] = [];
  const append = (values: readonly number[]): void => {
    for (const id of values) ids.push(id);
  };
  for (const table of tablesIn(document, story)) {
    append(propertyIds(table.propertyChanges));
    if (table.formatting?.gridChange !== undefined) ids.push(table.formatting.gridChange.id);
    for (const row of table.rows) {
      append(propertyIds(row.propertyChanges));
      append(propertyIds(row.tablePropertyExceptionChanges));
      if (row.structuralChange !== undefined) ids.push(row.structuralChange.info.id);
      for (const cell of row.cells) {
        append(propertyIds(cell.propertyChanges));
        append(cellStructuralIds(cell));
      }
    }
  }
  return ids;
};

type PropertyChangeLike<Formatting> = {
  info: { id: number };
  previousFormatting?: Formatting;
};
type PropertyResolution<Change, Formatting> = {
  changes: Change[] | undefined;
  formatting: Formatting | undefined;
  restoreFormatting: boolean;
  changed: boolean;
};
type ResolvePropertiesOptions<Change, Formatting> = {
  changes: readonly Change[] | undefined;
  formatting: Formatting | undefined;
  ids: ReadonlySet<number>;
  reject: boolean;
};

/** Removing selected changes rebases the next retained snapshot, or restores the last removed snapshot. */
const resolveProperties = <Change extends PropertyChangeLike<Formatting>, Formatting>({
  changes,
  formatting,
  ids,
  reject,
}: ResolvePropertiesOptions<Change, Formatting>): PropertyResolution<Change, Formatting> => {
  if (changes === undefined || !changes.some(({ info }) => ids.has(info.id))) {
    return {
      changes: changes === undefined ? undefined : [...changes],
      formatting,
      restoreFormatting: false,
      changed: false,
    };
  }
  if (!reject) {
    const remaining = changes.filter(({ info }) => !ids.has(info.id));
    return {
      changes: remaining.length === 0 ? undefined : remaining,
      formatting,
      restoreFormatting: false,
      changed: true,
    };
  }
  const remaining: Change[] = [];
  let previous: Formatting | undefined = undefined;
  let removing = false;
  for (const change of changes) {
    if (ids.has(change.info.id)) {
      if (!removing) previous = change.previousFormatting;
      removing = true;
      continue;
    }
    if (!removing) {
      remaining.push(change);
      continue;
    }
    const rebased = { ...change };
    delete rebased.previousFormatting;
    if (previous !== undefined) rebased.previousFormatting = previous;
    remaining.push(rebased);
    previous = undefined;
    removing = false;
  }
  return {
    changes: remaining.length === 0 ? undefined : remaining,
    formatting: removing ? previous : formatting,
    restoreFormatting: removing,
    changed: true,
  };
};

type PropertyTableOptions = { table: Table; ids: ReadonlySet<number>; reject: boolean };
const propertyTable = ({
  table,
  ids,
  reject,
}: PropertyTableOptions): Result<Table, "unsupported"> => {
  const gridSnapshot = table.formatting?.gridChange;
  const selectedGrid = gridSnapshot !== undefined && ids.has(gridSnapshot.id);
  const tableSnapshot = table.propertyChanges?.at(-1);
  const selectedTableSnapshot =
    selectedGrid && tableSnapshot !== undefined && ids.has(tableSnapshot.info.id)
      ? tableSnapshot
      : undefined;
  if (
    gridSnapshot !== undefined &&
    ids.has(gridSnapshot.id) &&
    reject &&
    gridSnapshot.columnWidths.some((width) => width === undefined)
  ) {
    return Result.err("unsupported");
  }
  let next = table;
  const properties = resolveProperties({
    changes: table.propertyChanges,
    formatting: table.formatting,
    ids,
    reject,
  });
  if (properties.changed) {
    next = { ...next };
    if (properties.changes === undefined) delete next.propertyChanges;
    else next.propertyChanges = properties.changes;
    if (properties.restoreFormatting) {
      if (properties.formatting === undefined) delete next.formatting;
      else next.formatting = properties.formatting;
    }
  }
  if (reject && selectedTableSnapshot !== undefined && !properties.restoreFormatting) {
    if (selectedTableSnapshot.previousFormatting === undefined) delete next.formatting;
    else next.formatting = selectedTableSnapshot.previousFormatting;
  }
  if (gridSnapshot !== undefined && ids.has(gridSnapshot.id)) {
    const formatting = next.formatting;
    if (formatting?.gridChange?.id === gridSnapshot.id) {
      const withoutGrid = { ...formatting };
      delete withoutGrid.gridChange;
      if (
        Object.keys(withoutGrid).length === 0 &&
        (selectedTableSnapshot === undefined ||
          selectedTableSnapshot.previousFormatting === undefined)
      ) {
        delete next.formatting;
      } else {
        next.formatting = withoutGrid;
      }
    }
    if (reject) {
      next = {
        ...next,
        columnWidths: gridSnapshot.columnWidths.flatMap((width) =>
          width === undefined ? [] : [width],
        ),
      };
    }
  }
  return Result.ok(next);
};

type PropertyRowOptions = { row: TableRow; ids: ReadonlySet<number>; reject: boolean };
const propertyRow = ({ row, ids, reject }: PropertyRowOptions): TableRow => {
  let next = row;
  const rowProperties = resolveProperties({
    changes: row.propertyChanges,
    formatting: row.formatting,
    ids,
    reject,
  });
  if (rowProperties.changed) {
    next = { ...next };
    if (rowProperties.changes === undefined) delete next.propertyChanges;
    else next.propertyChanges = rowProperties.changes;
    if (rowProperties.restoreFormatting) {
      if (rowProperties.formatting === undefined) delete next.formatting;
      else next.formatting = rowProperties.formatting;
    }
  }
  const exceptionChanges = resolveProperties({
    changes: row.tablePropertyExceptionChanges,
    formatting: row.tablePropertyExceptions,
    ids,
    reject,
  });
  if (exceptionChanges.changed) {
    next = { ...next };
    if (exceptionChanges.changes === undefined) delete next.tablePropertyExceptionChanges;
    else next.tablePropertyExceptionChanges = exceptionChanges.changes;
    if (exceptionChanges.restoreFormatting) {
      if (exceptionChanges.formatting === undefined) delete next.tablePropertyExceptions;
      else next.tablePropertyExceptions = exceptionChanges.formatting;
    }
  }
  return next;
};

type PropertyCellOptions = { cell: TableCell; ids: ReadonlySet<number>; reject: boolean };
const propertyCell = ({ cell, ids, reject }: PropertyCellOptions): TableCell => {
  const changes = cell.propertyChanges;
  if (changes === undefined || !changes.some(({ info }) => ids.has(info.id))) return cell;
  const next = { ...cell };
  if (!reject) {
    const remaining = changes.filter(({ info }) => !ids.has(info.id));
    if (remaining.length === 0) delete next.propertyChanges;
    else next.propertyChanges = remaining;
    return next;
  }
  const remaining: TableCellPropertyChange[] = [];
  const mergeChange =
    cell.structuralChange?.type === "tableCellMerge" ? cell.structuralChange : undefined;
  const preserveMergeState =
    mergeChange !== undefined && !ids.has(mergeChange.info.id)
      ? cell.formatting?.vMerge
      : undefined;
  let priorFormatting: TableCellPropertyChange["previousFormatting"];
  let priorStructuralChange: TableCellPropertyChange["previousStructuralChange"];
  let removing = false;
  for (const change of changes) {
    if (ids.has(change.info.id)) {
      if (!removing) {
        priorFormatting = change.previousFormatting;
        priorStructuralChange = change.previousStructuralChange;
      }
      removing = true;
      continue;
    }
    if (!removing) {
      remaining.push(change);
      continue;
    }
    const rebased = { ...change };
    delete rebased.previousFormatting;
    delete rebased.previousStructuralChange;
    if (priorFormatting !== undefined) rebased.previousFormatting = priorFormatting;
    if (priorStructuralChange !== undefined)
      rebased.previousStructuralChange = priorStructuralChange;
    remaining.push(rebased);
    priorFormatting = undefined;
    priorStructuralChange = undefined;
    removing = false;
  }
  if (remaining.length === 0) delete next.propertyChanges;
  else next.propertyChanges = remaining;
  if (removing) {
    if (priorFormatting === undefined) delete next.formatting;
    else next.formatting = priorFormatting;
    if (preserveMergeState !== undefined) {
      next.formatting = { ...next.formatting, vMerge: preserveMergeState };
    }
    if (priorStructuralChange !== undefined) next.structuralChange = priorStructuralChange;
    else if (mergeChange !== undefined && ids.has(mergeChange.info.id))
      delete next.structuralChange;
  }
  return next;
};

type GridCell = { cell: TableCell; start: number; end: number };
type GridRow = { cells: GridCell[]; before: number; after: number };

/** Remove tracked inserted/deleted cells only when their selected cells describe complete columns. */
type ResolveCellColumnsOptions = { table: Table; ids: ReadonlySet<number>; reject: boolean };
const resolveCellColumns = ({
  table,
  ids,
  reject,
}: ResolveCellColumnsOptions): Result<Table, "unsupported"> => {
  const columnWidths = table.columnWidths;
  if (columnWidths === undefined || columnWidths.length === 0) return Result.err("unsupported");
  const width = columnWidths.length;
  const removeCells = (change: TableStructuralChangeInfo): boolean => {
    if (change.type === "tableCellInsertion") return reject;
    if (change.type === "tableCellDeletion") return !reject;
    return false;
  };
  const gridRows: GridRow[] = [];
  const targetSlots = new Set<number>();
  for (const row of table.rows) {
    const before = row.formatting?.gridBefore ?? 0;
    const after = row.formatting?.gridAfter ?? 0;
    if (!Number.isInteger(before) || !Number.isInteger(after) || before < 0 || after < 0) {
      return Result.err("unsupported");
    }
    let cursor = before;
    const cells: GridCell[] = [];
    for (const cell of row.cells) {
      const span = cell.formatting?.gridSpan ?? 1;
      if (!Number.isInteger(span) || span < 1) return Result.err("unsupported");
      const end = cursor + span;
      if (end > width - after) return Result.err("unsupported");
      const change = cell.structuralChange;
      if (change !== undefined && ids.has(change.info.id) && removeCells(change)) {
        for (let slot = cursor; slot < end; slot++) targetSlots.add(slot);
      }
      cells.push({ cell, start: cursor, end });
      cursor = end;
    }
    if (cursor > width - after) return Result.err("unsupported");
    gridRows.push({ cells, before, after });
  }
  if (targetSlots.size === 0) return Result.err("unsupported");

  const removedByRow: Set<number>[] = [];
  for (const row of gridRows) {
    const removed = new Set<number>();
    for (const { cell, start, end } of row.cells) {
      const overlaps = Array.from({ length: end - start }, (_, offset) =>
        targetSlots.has(start + offset),
      );
      const isSelectedRemoval =
        cell.structuralChange !== undefined &&
        ids.has(cell.structuralChange.info.id) &&
        removeCells(cell.structuralChange);
      if (
        overlaps.some(Boolean) !== isSelectedRemoval ||
        (overlaps.some(Boolean) && !overlaps.every(Boolean))
      ) {
        return Result.err("unsupported");
      }
      if (isSelectedRemoval) for (let slot = start; slot < end; slot++) removed.add(slot);
    }
    for (let slot = 0; slot < row.before; slot++) if (targetSlots.has(slot)) removed.add(slot);
    for (let slot = width - row.after; slot < width; slot++)
      if (targetSlots.has(slot)) removed.add(slot);
    if (removed.size !== targetSlots.size) return Result.err("unsupported");
    removedByRow.push(removed);
  }

  const nextWidths = columnWidths.filter((_columnWidth, index) => !targetSlots.has(index));
  if (nextWidths.length === 0) return Result.err("unsupported");
  const rows = table.rows.map((row, rowIndex) => {
    const layout = gridRows[rowIndex];
    const removed = removedByRow[rowIndex];
    if (layout === undefined || removed === undefined) return row;
    const nextRow = { ...row };
    nextRow.cells = layout.cells.flatMap(({ cell }) => {
      const change = cell.structuralChange;
      if (change !== undefined && ids.has(change.info.id) && removeCells(change)) return [];
      if (change !== undefined && ids.has(change.info.id) && change.type !== "tableCellMerge") {
        const nextCell = { ...cell };
        delete nextCell.structuralChange;
        return [nextCell];
      }
      return [cell];
    });
    const removedBefore = Array.from({ length: layout.before }, (_, slot) =>
      removed.has(slot),
    ).filter(Boolean).length;
    const removedAfter = Array.from({ length: layout.after }, (_, offset) =>
      removed.has(width - layout.after + offset),
    ).filter(Boolean).length;
    if (removedBefore > 0 || removedAfter > 0) {
      const formatting = { ...row.formatting };
      const before = layout.before - removedBefore;
      const after = layout.after - removedAfter;
      if (before === 0) delete formatting.gridBefore;
      else formatting.gridBefore = before;
      if (after === 0) delete formatting.gridAfter;
      else formatting.gridAfter = after;
      if (Object.keys(formatting).length === 0) delete nextRow.formatting;
      else nextRow.formatting = formatting;
    }
    return nextRow;
  });
  return Result.ok({ ...table, rows, columnWidths: nextWidths });
};

type ResolveCellChangesOptions = {
  table: Table;
  ids: ReadonlySet<number>;
  reject: boolean;
  selectedGridChange: boolean;
};
const resolveCellChanges = ({
  table,
  ids,
  reject,
  selectedGridChange,
}: ResolveCellChangesOptions): Result<Table, "unsupported"> => {
  const properties = (source: Table): Table => ({
    ...source,
    rows: source.rows.map((row) => ({
      ...propertyRow({ row, ids, reject }),
      cells: row.cells.map((cell) => propertyCell({ cell, ids, reject })),
    })),
  });
  const resolveStructure = (source: Table): Table => ({
    ...source,
    rows: source.rows.map((row) => ({
      ...row,
      cells: row.cells.flatMap((cell) => {
        const change = cell.structuralChange;
        if (change === undefined || !ids.has(change.info.id)) return [cell];
        if (change.type === "tableCellMerge") {
          const copy = { ...cell };
          if (reject) {
            const formatting = { ...copy.formatting };
            if (change.verticalMergeOriginal === undefined) delete formatting.vMerge;
            else
              formatting.vMerge = change.verticalMergeOriginal === "rest" ? "restart" : "continue";
            if (Object.keys(formatting).length === 0) delete copy.formatting;
            else copy.formatting = formatting;
          }
          if (copy.structuralChange === change) delete copy.structuralChange;
          return [copy];
        }
        if (change.type === "tableCellInsertion" || change.type === "tableCellDeletion") {
          if ((change.type === "tableCellInsertion") === reject) return [];
          const copy = { ...cell };
          if (copy.structuralChange === change) delete copy.structuralChange;
          return [copy];
        }
        return [cell];
      }),
    })),
  });

  const propertiesResolved = properties(table);
  const structureResolved = resolveStructure(propertiesResolved);
  if (selectedGridChange) return Result.ok(structureResolved);

  const hasRemovedCells = table.rows.some((row) =>
    row.cells.some((cell) => {
      const change = cell.structuralChange;
      return (
        change !== undefined &&
        ids.has(change.info.id) &&
        ((change.type === "tableCellInsertion" && reject) ||
          (change.type === "tableCellDeletion" && !reject))
      );
    }),
  );
  if (!hasRemovedCells) return Result.ok(structureResolved);

  // Cell insertions used by horizontal/vertical splits restore the owner's span
  // or merge state. Keep the grid when that restored topology covers it exactly.
  if (tableGrid(structureResolved, DOCUMENT_OP_TYPES.SET_TABLE).isOk()) {
    return Result.ok(structureResolved);
  }

  // Historical cell marks without a grid snapshot can represent complete
  // column edits. This fallback removes those slots from the explicit grid.
  const columns = resolveCellColumns({ table: propertiesResolved, ids, reject });
  if (columns.isErr()) return columns;
  return Result.ok(resolveStructure(columns.value));
};

type ResolveTableRowsOptions = { document: Document; op: ResolveRevisionOp; applyOps: ApplyOps };
const refusal = (op: ResolveRevisionOp, message: string) =>
  new DocumentOpRefusal({
    opType: op.type,
    reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    message,
  });

export const resolveTableRows = ({
  document,
  op,
  applyOps,
}: ResolveTableRowsOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  const ids = new Set(op.revisionIds);
  let current = document;
  const edits: DocumentEdit[] = [];

  // Keep row insertion/deletion on its established primitive, which also
  // removes an emptied table. Table property and cell changes use setTable.
  for (const originalTable of tablesIn(document, op.story)) {
    if (
      !originalTable.rows.some(
        (row) => row.structuralChange !== undefined && ids.has(row.structuralChange.info.id),
      )
    )
      continue;
    const blockId = tableRowAnchor(originalTable.rows);
    if (blockId === undefined)
      return Result.err(refusal(op, "The table has no paragraph that can address its rows."));
    const target = {
      type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
      story: op.story,
      blockId,
      expected: originalTable.rows,
      rows: originalTable.rows,
    } as const;
    const located = locateTableRow(current, target);
    if (located.isErr()) return Result.err(located.error);
    const rows = located.value.table.rows.flatMap((row) => {
      const change = row.structuralChange;
      if (change === undefined || !ids.has(change.info.id)) return [row];
      const added = change.type === "tableRowInsertion";
      if (added !== (op.decision === REVISION_DECISIONS.ACCEPT)) return [];
      const next = { ...row };
      delete next.structuralChange;
      return [next];
    });
    const applied = applyOps(current, [
      {
        type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
        story: op.story,
        blockId,
        expected: located.value.table.rows,
        rows,
      },
    ]);
    if (applied.isErr()) return Result.err(applied.error);
    edits.push(applied.value);
    current = applied.value.document;
  }

  for (const originalTable of tablesIn(current, op.story)) {
    const originalIds = reachableOnTable(originalTable);
    if (![...originalIds].some((id) => ids.has(id))) continue;
    const blockId = tableRowAnchor(originalTable.rows);
    if (blockId === undefined)
      return Result.err(
        refusal(op, "The table has no paragraph that can address its review records."),
      );
    const target = { type: DOCUMENT_OP_TYPES.SET_TABLE, story: op.story, blockId } as const;
    const located = locateTableRow(current, target);
    if (located.isErr()) return Result.err(located.error);
    const table = located.value.table;
    const reject = op.decision === REVISION_DECISIONS.REJECT;
    const selectedGridChange =
      table.formatting?.gridChange !== undefined && ids.has(table.formatting.gridChange.id);
    const cellsResolved = resolveCellChanges({ table, ids, reject, selectedGridChange });
    if (cellsResolved.isErr())
      return Result.err(
        refusal(op, "The selected cell revisions do not describe complete grid columns."),
      );
    const propertiesResolved = propertyTable({ table: cellsResolved.value, ids, reject });
    if (propertiesResolved.isErr())
      return Result.err(
        refusal(op, "The previous grid contains widths the table model cannot restore."),
      );
    const nextTable = propertiesResolved.value;
    const applied = applyOps(current, [
      {
        type: DOCUMENT_OP_TYPES.SET_TABLE,
        story: op.story,
        blockId,
        expected: table,
        table: nextTable,
      },
    ]);
    if (applied.isErr()) return Result.err(applied.error);
    edits.push(applied.value);
    current = applied.value.document;
  }
  return Result.ok(combineEdits(document, edits));
};

const reachableOnTable = (table: Table): Set<number> =>
  new Set([
    ...propertyIds(table.propertyChanges),
    ...(table.formatting?.gridChange === undefined ? [] : [table.formatting.gridChange.id]),
    ...table.rows.flatMap((row) => [
      ...propertyIds(row.propertyChanges),
      ...propertyIds(row.tablePropertyExceptionChanges),
      ...(row.structuralChange === undefined ? [] : [row.structuralChange.info.id]),
      ...row.cells.flatMap((cell) => [
        ...propertyIds(cell.propertyChanges),
        ...cellStructuralIds(cell),
      ]),
    ]),
  ]);
