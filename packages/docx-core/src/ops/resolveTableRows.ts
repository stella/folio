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
import { maxOf } from "../model/extremes";
import {
  declaredUnmeasuredColumnCount,
  statedTableColumnWidths,
  withTableColumnWidths,
} from "../model/tableColumnWidths";
import { storyBody, storyParagraphs } from "./blocks";
import { combineEdits, type DocumentEdit } from "./edits";
import { IDENTITY_SPACES, identityKeysIn, slotKey } from "./ids";
import {
  pendingMergedOwner,
  relocatePositionedMarkers,
} from "./positionedMarkers";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { ApplyOps } from "./resolve";
import { reviewFieldsOf } from "./review";
import {
  resolveTableMoves,
  tableMoveOwnershipMatches,
  tableMoveRevisionIds,
} from "./tableCellMove";
import { editRemovesRow, tableGrid } from "./tableGrid";
import { locateTableRow, tableRowAnchor } from "./tableLocation";
import {
  removalOrphansAMerge,
  rowsWithoutRow,
  withVerticalMerge,
} from "./tableRowRemoval";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type OpStory,
  type ResolveRevisionOp,
  REVISION_DECISIONS,
} from "./types";

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
      case "commentRangeStart":
      case "commentRangeEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
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

/** A geometry decision removes or retains the cells and the grid together. */
export const withBelongingTableGeometry = (
  document: Document,
  op: ResolveRevisionOp,
): Result<ResolveRevisionOp, DocumentOpRefusal> => {
  const ids = new Set(op.revisionIds);
  const uniqueCount = ids.size;
  for (const table of tablesIn(document, op.story)) {
    const grid = table.formatting?.gridChange;
    const geometry = grid?.geometryChange;
    const members = geometry?.revisionIds;
    if (grid === undefined || !Array.isArray(members)) continue;
    const local = reachableOnTable(table);
    const moves =
      geometry?.type === "mergeCells"
        ? [
            ...(Array.isArray(geometry.moves)
              ? tableMoveRevisionIds(geometry.moves)
              : []),
            ...(Array.isArray(geometry.nestedMoves)
              ? geometry.nestedMoves.flatMap((move) => move.rows)
              : []),
          ]
        : [];
    const selectsGeometry =
      ids.has(grid.id) || members.some((id) => local.has(id) && ids.has(id));
    if (!selectsGeometry && !moves.some((id) => ids.has(id))) continue;
    if (
      geometry?.type === "mergeCells" &&
      !tableMoveOwnershipMatches(table, geometry)
    ) {
      if (!selectsGeometry) {
        const physical = new Set(identityKeysIn(table));
        if (
          !moves.some(
            (id) =>
              ids.has(id) &&
              physical.has(slotKey({ space: IDENTITY_SPACES.REVISION, id })),
          )
        )
          continue;
      }
      return Result.err(
        new DocumentOpRefusal({
          opType: op.type,
          reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
          message:
            "The tracked table merge no longer matches its native move records.",
        }),
      );
    }
    ids.add(grid.id);
    for (const id of members) if (local.has(id)) ids.add(id);
    // The topology decision unwraps these paragraphs and hands back their
    // source ids. An independent inline decision would lose the handoff.
    for (const id of moves) ids.delete(id);
  }
  return Result.ok(
    ids.size === uniqueCount && op.revisionIds.every((id) => ids.has(id))
      ? op
      : { ...op, revisionIds: [...ids] },
  );
};

const propertyIds = (
  changes: readonly { info: { id: number } }[] | undefined,
): number[] => changes?.map(({ info }) => info.id) ?? [];
const cellStructuralIds = (cell: TableCell): number[] =>
  cell.structuralChange === undefined ? [] : [cell.structuralChange.info.id];

/** Revision ids the table resolver can reach, including properties and cell marks. */
export const reachableRowIds = (
  document: Document,
  story: OpStory,
): number[] => {
  const ids: number[] = [];
  const append = (values: readonly number[]): void => {
    for (const id of values) ids.push(id);
  };
  for (const table of tablesIn(document, story)) {
    append(propertyIds(table.propertyChanges));
    if (table.formatting?.gridChange !== undefined)
      ids.push(table.formatting.gridChange.id);
    for (const row of table.rows) {
      append(propertyIds(row.propertyChanges));
      append(propertyIds(row.tablePropertyExceptionChanges));
      if (row.structuralChange !== undefined)
        ids.push(row.structuralChange.info.id);
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
const resolveProperties = <
  Change extends PropertyChangeLike<Formatting>,
  Formatting,
>({
  changes,
  formatting,
  ids,
  reject,
}: ResolvePropertiesOptions<Change, Formatting>): PropertyResolution<
  Change,
  Formatting
> => {
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

type PropertyTableOptions = {
  table: Table;
  ids: ReadonlySet<number>;
  reject: boolean;
  /** Rows still hold columns a rejected grid change added: its grid stays. */
  keepsGrid?: boolean;
};
/** The grid columns the first row covers, or undefined when the rows differ. */
const coveredBy = (row: TableRow): number =>
  (row.formatting?.gridBefore ?? 0) +
  (row.formatting?.gridAfter ?? 0) +
  row.cells.reduce((sum, cell) => sum + (cell.formatting?.gridSpan ?? 1), 0);

const rowCoverage = (table: Table): number | undefined => {
  let coverage: number | undefined;
  for (const row of table.rows) {
    const covered = coveredBy(row);
    if (coverage !== undefined && covered !== coverage) return undefined;
    coverage = covered;
  }
  return coverage;
};

/** Every row covers the same number of grid columns. */
const coversEvenly = (table: Table): boolean =>
  rowCoverage(table) !== undefined;

/**
 * The rows imply a grid this wide: the widest covers it, and a row that ends
 * early leaves the columns after it.
 */
const impliesWidth = (table: Table, width: number): boolean =>
  table.rows.length > 0 &&
  maxOf(table.rows.map((row) => coveredBy(row))) === width;

const propertyTable = ({
  table,
  ids,
  reject,
  keepsGrid = false,
}: PropertyTableOptions): Table => {
  const gridSnapshot = table.formatting?.gridChange;
  const selectedGrid = gridSnapshot !== undefined && ids.has(gridSnapshot.id);
  const tableSnapshot = table.propertyChanges?.at(-1);
  const selectedTableSnapshot =
    selectedGrid &&
    tableSnapshot !== undefined &&
    ids.has(tableSnapshot.info.id)
      ? tableSnapshot
      : undefined;
  let next = selectedGrid ? { ...table } : table;
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
  if (
    reject &&
    selectedTableSnapshot !== undefined &&
    !properties.restoreFormatting
  ) {
    if (selectedTableSnapshot.previousFormatting === undefined)
      delete next.formatting;
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
    // The earlier grid comes back, unless a row keeps a cell in the columns
    // the change added: the columns it holds stay.
    if (reject && !keepsGrid)
      next = withTableColumnWidths(next, gridSnapshot.columnWidths);
  }
  return next;
};

type PropertyRowOptions = {
  row: TableRow;
  ids: ReadonlySet<number>;
  reject: boolean;
};
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
    if (exceptionChanges.changes === undefined)
      delete next.tablePropertyExceptionChanges;
    else next.tablePropertyExceptionChanges = exceptionChanges.changes;
    if (exceptionChanges.restoreFormatting) {
      if (exceptionChanges.formatting === undefined)
        delete next.tablePropertyExceptions;
      else next.tablePropertyExceptions = exceptionChanges.formatting;
    }
  }
  return next;
};

type PropertyCellOptions = {
  cell: TableCell;
  ids: ReadonlySet<number>;
  reject: boolean;
};
const propertyCell = ({
  cell,
  ids,
  reject,
}: PropertyCellOptions): TableCell => {
  const changes = cell.propertyChanges;
  if (changes === undefined || !changes.some(({ info }) => ids.has(info.id)))
    return cell;
  const next = { ...cell };
  if (!reject) {
    const remaining = changes.filter(({ info }) => !ids.has(info.id));
    if (remaining.length === 0) delete next.propertyChanges;
    else next.propertyChanges = remaining;
    return next;
  }
  const remaining: TableCellPropertyChange[] = [];
  const mergeChange =
    cell.structuralChange?.type === "tableCellMerge"
      ? cell.structuralChange
      : undefined;
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
    if (priorFormatting !== undefined)
      rebased.previousFormatting = priorFormatting;
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
    if (priorStructuralChange !== undefined)
      next.structuralChange = priorStructuralChange;
    else if (mergeChange !== undefined && ids.has(mergeChange.info.id))
      delete next.structuralChange;
  }
  return next;
};

type GridCell = { cell: TableCell; start: number; end: number };
type GridRow = {
  cells: GridCell[];
  before: number;
  cellsEnd: number;
  emptied: boolean;
};

/**
 * Remove the selected tracked inserted (rejected) or deleted (accepted) cells.
 * A grid column only goes when every row loses what covers it; a row's
 * `gridBefore` and `gridAfter` columns, and those a row that ends early
 * leaves after it, keep theirs. Where other rows keep a removed cell's
 * columns, the cell before it in its row (the one after it when it starts the
 * row) widens over them. A row that loses every cell it has, or every cell
 * it starts, goes, and keeps no column.
 */
type ResolveCellColumnsOptions = {
  table: Table;
  ids: ReadonlySet<number>;
  reject: boolean;
};
const resolveCellColumns = ({
  table,
  ids,
  reject,
}: ResolveCellColumnsOptions): Result<
  { table: Table; holdsRemovedColumns: boolean },
  "unsupported"
> => {
  const removeCells = (change: TableStructuralChangeInfo): boolean => {
    if (change.type === "tableCellInsertion") return reject;
    if (change.type === "tableCellDeletion") return !reject;
    return false;
  };
  const removes = (cell: TableCell): boolean =>
    cell.structuralChange !== undefined &&
    ids.has(cell.structuralChange.info.id) &&
    removeCells(cell.structuralChange);
  const columnWidths = statedTableColumnWidths(table);
  const gridRows: GridRow[] = [];
  const ends: number[] = [];
  const losesEveryCell = (row: TableRow) =>
    row.cells.length > 0 && row.cells.every(removes);
  for (const row of table.rows) {
    if (losesEveryCell(row)) {
      gridRows.push({ cells: [], before: 0, cellsEnd: 0, emptied: true });
      continue;
    }
    const before = row.formatting?.gridBefore ?? 0;
    const after = row.formatting?.gridAfter ?? 0;
    if (
      !Number.isInteger(before) ||
      !Number.isInteger(after) ||
      before < 0 ||
      after < 0
    ) {
      return Result.err("unsupported");
    }
    let cursor = before;
    const cells: GridCell[] = [];
    for (const cell of row.cells) {
      const span = cell.formatting?.gridSpan ?? 1;
      if (!Number.isInteger(span) || span < 1) return Result.err("unsupported");
      cells.push({ cell, start: cursor, end: cursor + span });
      cursor += span;
    }
    gridRows.push({ cells, before, cellsEnd: cursor, emptied: false });
    ends.push(cursor + after);
  }
  const gridWidth = maxOf(ends, columnWidths?.length ?? 0);
  if (gridWidth === 0) return Result.err("unsupported");

  // A slot goes when no row keeps anything in it.
  const removable = new Set<number>();
  for (let slot = 0; slot < gridWidth; slot++) {
    const kept = gridRows.some(
      ({ cells, before, cellsEnd, emptied }) =>
        !emptied &&
        (slot < before ||
          slot >= cellsEnd ||
          cells.some(
            ({ cell, start, end: cellEnd }) =>
              slot >= start && slot < cellEnd && !removes(cell),
          )),
    );
    if (!kept) removable.add(slot);
  }
  const removedIn = (start: number, end: number) =>
    Array.from({ length: end - start }, (_, offset) =>
      removable.has(start + offset),
    ).filter(Boolean).length;

  // Whether a removed cell left columns other rows still hold.
  let holdsRemovedColumns = false;
  const rows: TableRow[] = [];
  for (const [rowIndex, row] of table.rows.entries()) {
    const layout = gridRows[rowIndex];
    if (layout === undefined) return Result.err("unsupported");
    if (layout.emptied) continue;
    const cells: TableCell[] = [];
    // Columns still occupied in other rows can move to the following cell
    // when the previous cell cannot grow without changing its merge.
    let carried = 0;
    let carriedFrom: TableCell | undefined;
    let insertedPlaceholder = false;
    let omittedBefore = 0;
    let omittedAfter = 0;
    const lastKept = layout.cells.findLastIndex(({ cell }) => !removes(cell));
    const widen = (cell: TableCell, extra: number): TableCell | undefined => {
      if (extra === 0) return cell;
      // A widened cell would no longer line up with its vertical merge.
      if (cell.formatting?.vMerge !== undefined) return undefined;
      return {
        ...cell,
        formatting: {
          ...cell.formatting,
          gridSpan: (cell.formatting?.gridSpan ?? 1) + extra,
        },
      };
    };
    for (const [index, { cell, start, end }] of layout.cells.entries()) {
      if (removes(cell)) {
        const kept = end - start - removedIn(start, end);
        if (kept === 0) continue;
        holdsRemovedColumns = true;
        const previous = cells.at(-1);
        if (previous === undefined) {
          carried += kept;
          continue;
        }
        const widened = widen(previous, kept);
        if (widened === undefined) {
          if (index < lastKept) {
            carried += kept;
            carriedFrom ??= cell;
          } else omittedAfter += kept;
          continue;
        }
        cells[cells.length - 1] = widened;
        continue;
      }
      let next = cell;
      const change = cell.structuralChange;
      if (
        change !== undefined &&
        ids.has(change.info.id) &&
        change.type !== "tableCellMerge"
      ) {
        next = { ...cell };
        delete next.structuralChange;
      }
      // A kept cell holds no column every row loses.
      const widened = widen(next, carried);
      if (widened === undefined) {
        if (cells.length === 0) {
          omittedBefore = carried;
        } else {
          // An interior slot between vertical merges needs a real cell.
          // Its original paragraph's address remains on an empty terminator.
          const paraId =
            carriedFrom === undefined
              ? undefined
              : tableRowAnchor([{ type: "tableRow", cells: [carriedFrom] }]);
          if (paraId === undefined) return Result.err("unsupported");
          cells.push({
            type: "tableCell",
            content: [{ type: "paragraph", paraId, content: [] }],
            ...(carried === 1 ? {} : { formatting: { gridSpan: carried } }),
          });
          insertedPlaceholder = true;
        }
      }
      carried = 0;
      carriedFrom = undefined;
      cells.push(widened ?? next);
    }
    if (carried > 0) return Result.err("unsupported");
    // Accepting a deletion that leaves the row nothing it starts removes
    // the row, as the direct deletion does.
    if (!reject) {
      const survivors = insertedPlaceholder
        ? row.cells.filter((cell) => !removes(cell))
        : cells;
      if (editRemovesRow({ before: row.cells, after: survivors })) continue;
    }
    // Preserve the occupied columns of a vertical merge. Unoccupied slots
    // at the row edges can be expressed as native gridBefore/gridAfter.
    rows.push({
      ...row,
      cells,
      ...(omittedBefore === 0 && omittedAfter === 0
        ? {}
        : {
            formatting: {
              ...row.formatting,
              ...(omittedBefore === 0
                ? {}
                : { gridBefore: layout.before + omittedBefore }),
              ...(omittedAfter === 0
                ? {}
                : {
                    gridAfter: (row.formatting?.gridAfter ?? 0) + omittedAfter,
                  }),
            },
          }),
    });
  }
  if (rows.length === 0) return Result.err("unsupported");
  let next: Table = { ...table, rows };
  if (columnWidths !== undefined && columnWidths.length > 0) {
    const nextWidths = columnWidths.filter(
      (_width, slot) => !removable.has(slot),
    );
    // A document can extend its declared grid with unmeasured cell spans.
    next = withTableColumnWidths(next, nextWidths);
  }
  return Result.ok({ table: next, holdsRemovedColumns });
};

const selectedMergeMembers = (table: Table, ids: ReadonlySet<number>) => {
  const grid = table.formatting?.gridChange;
  return grid?.geometryChange?.type === "mergeCells" && ids.has(grid.id)
    ? new Set(grid.geometryChange.revisionIds)
    : undefined;
};

const mergeClosesRow = (
  row: TableRow,
  selected: ReadonlySet<number> | undefined,
): boolean => {
  const change = row.structuralChange;
  return (
    change?.type === "tableRowDeletion" &&
    selected?.has(change.info.id) === true
  );
};

type EndAdjacentMergesOptions = {
  before: Table;
  after: Table;
  selected: ReadonlySet<number> | undefined;
};
/** End a neighboring vertical merge that the closed rows alone continued. */
const endAdjacentMerges = ({
  before,
  after,
  selected,
}: EndAdjacentMergesOptions): Table => {
  if (
    selected === undefined ||
    !before.rows.some((row) => mergeClosesRow(row, selected))
  )
    return after;
  const previous = tableGrid(before, DOCUMENT_OP_TYPES.SET_TABLE);
  if (previous.isErr()) return after;
  const candidates = new Set<string>();
  for (const [index, row] of before.rows.entries()) {
    if (!mergeClosesRow(row, selected)) continue;
    for (const slot of previous.value.rows[index] ?? []) {
      if (slot.ownerRow === index) continue;
      const owner = before.rows[slot.ownerRow]?.cells[slot.ownerIndex];
      if (
        owner === undefined ||
        owner.propertyChanges?.some(({ info }) => selected.has(info.id))
      )
        continue;
      const anchor = tableRowAnchor([{ type: "tableRow", cells: [owner] }]);
      if (anchor !== undefined) candidates.add(anchor);
    }
  }
  if (candidates.size === 0) return after;
  // The pending snapshot interprets the original geometry. The surviving rows
  // already hold the newly accepted geometry.
  const { gridChange: _before, ...formatting } = after.formatting ?? {};
  const current = tableGrid(
    { ...after, formatting },
    DOCUMENT_OP_TYPES.SET_TABLE,
  );
  if (current.isErr()) return after;
  const ended = new Set<TableCell>();
  for (const [index, slots] of current.value.rows.entries()) {
    for (const slot of slots) {
      if (slot.cell.formatting?.vMerge !== "restart") continue;
      const anchor = tableRowAnchor([{ type: "tableRow", cells: [slot.cell] }]);
      if (anchor === undefined || !candidates.has(anchor)) continue;
      if (
        !current.value.rows[index + 1]?.some(
          ({ ownerRow, ownerIndex }) =>
            ownerRow === index && ownerIndex === slot.index,
        )
      )
        ended.add(slot.cell);
    }
  }
  if (ended.size === 0) return after;
  return {
    ...after,
    rows: after.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) =>
        ended.has(cell) ? withVerticalMerge(cell, undefined) : cell,
      ),
    })),
  };
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
}: ResolveCellChangesOptions): Result<
  { table: Table; holdsRemovedColumns: boolean },
  "unsupported"
> => {
  const mergeRows = reject ? undefined : selectedMergeMembers(table, ids);
  const properties = (source: Table): Table => ({
    ...source,
    rows: source.rows.map((row) => ({
      ...propertyRow({ row, ids, reject }),
      cells: row.cells.map((cell) => propertyCell({ cell, ids, reject })),
    })),
  });
  // A row whose every cell goes, goes with them.
  const resolveStructure = (source: Table): Table => ({
    ...source,
    rows: source.rows.flatMap((row) => {
      if (mergeClosesRow(row, mergeRows)) return [];
      const cells = row.cells.flatMap((cell) => {
        const change = cell.structuralChange;
        if (change === undefined || !ids.has(change.info.id)) return [cell];
        if (change.type === "tableCellMerge") {
          const copy = { ...cell };
          if (reject) {
            const formatting = { ...copy.formatting };
            if (change.verticalMergeOriginal === undefined)
              delete formatting.vMerge;
            else
              formatting.vMerge =
                change.verticalMergeOriginal === "rest"
                  ? "restart"
                  : "continue";
            if (Object.keys(formatting).length === 0) delete copy.formatting;
            else copy.formatting = formatting;
          }
          if (copy.structuralChange === change) delete copy.structuralChange;
          return [copy];
        }
        if (
          change.type === "tableCellInsertion" ||
          change.type === "tableCellDeletion"
        ) {
          if ((change.type === "tableCellInsertion") === reject) return [];
          const copy = { ...cell };
          if (copy.structuralChange === change) delete copy.structuralChange;
          return [copy];
        }
        return [cell];
      });
      // Accepting a deletion that leaves a row nothing it starts removes
      // the row; rejecting gives back the rows the source had.
      return (row.cells.length > 0 && cells.length === 0) ||
        (!reject && editRemovesRow({ before: row.cells, after: cells }))
        ? []
        : [{ ...row, cells }];
    }),
  });

  const propertiesResolved = properties(table);
  const structureResolved = resolveStructure(propertiesResolved);

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
  if (
    selectedGridChange &&
    table.formatting?.gridChange?.geometryChange !== undefined
  )
    return Result.ok({
      table: resolveTableMoves(
        endAdjacentMerges({
          before: table,
          after: structureResolved,
          selected: mergeRows,
        }),
        reject,
      ),
      holdsRemovedColumns: false,
    });
  if (!hasRemovedCells)
    return Result.ok({ table: structureResolved, holdsRemovedColumns: false });

  // Cell insertions used by horizontal/vertical splits restore the owner's span
  // or merge state. Keep the grid when that restored topology covers it
  // exactly, or when the grid the selection restores will. Removing cells
  // never narrows the grid by itself: a row that then ends early is resolved
  // by the column rule below.
  const restored = tableGrid(structureResolved, DOCUMENT_OP_TYPES.SET_TABLE);
  if (
    (restored.isOk() &&
      restored.value.stated &&
      restored.value.width ===
        tableGrid(table, DOCUMENT_OP_TYPES.SET_TABLE)
          .map(({ width }) => width)
          .unwrapOr(restored.value.width)) ||
    (selectedGridChange &&
      (coversEvenly(structureResolved) ||
        impliesWidth(
          structureResolved,
          (reject
            ? table.formatting?.gridChange?.columnWidths.length
            : (table.columnWidths?.length ??
              declaredUnmeasuredColumnCount(table.unmeasuredColumnCount))) ?? 0,
        )))
  ) {
    return Result.ok({ table: structureResolved, holdsRemovedColumns: false });
  }

  // Historical cell marks without a grid snapshot can represent complete
  // column edits. This fallback removes those slots from the explicit grid.
  const columns = resolveCellColumns({
    table: propertiesResolved,
    ids,
    reject,
  });
  if (columns.isErr()) return columns;
  return Result.ok({
    table: resolveStructure(columns.value.table),
    holdsRemovedColumns: columns.value.holdsRemovedColumns,
  });
};

type RemovesEveryCellOptions = {
  table: Table;
  ids: ReadonlySet<number>;
  reject: boolean;
};
/** Whether the decision removes every cell of the table: inserted ones rejected, deleted ones accepted. */
const removesEveryCell = ({ table, ids, reject }: RemovesEveryCellOptions) =>
  table.rows.every((row) =>
    row.cells.every((cell) => {
      const change = cell.structuralChange;
      return (
        change !== undefined &&
        ids.has(change.info.id) &&
        ((change.type === "tableCellInsertion" && reject) ||
          (change.type === "tableCellDeletion" && !reject))
      );
    }),
  ) && table.rows.some((row) => row.cells.length > 0);

type ResolveTableRowsOptions = {
  document: Document;
  op: ResolveRevisionOp;
  applyOps: ApplyOps;
};
const refusal = (op: ResolveRevisionOp, message: string) =>
  new DocumentOpRefusal({
    opType: op.type,
    reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    message,
  });

/**
 * The paragraph marks still pending in rows and cells this decision removes,
 * resolved away first: the paragraphs go with their row or cell, so no mark
 * of theirs is left pending in a structure that goes, which the inverse
 * restoring it could not state (a cell's final mark would be pending there).
 */
const marksInRemovedStructure = (
  document: Document,
  op: ResolveRevisionOp,
): DocumentOp[] => {
  const ids = new Set(op.revisionIds);
  const reject = op.decision === REVISION_DECISIONS.REJECT;
  const removes = (change: TableStructuralChangeInfo | undefined): boolean => {
    if (change === undefined || !ids.has(change.info.id)) return false;
    if (reject)
      return (
        change.type === "tableRowInsertion" ||
        change.type === "tableCellInsertion"
      );
    return (
      change.type === "tableRowDeletion" || change.type === "tableCellDeletion"
    );
  };
  // A merged cell whose row goes is handed on to the next row its merge
  // continues in that stays (`rowsWithoutRow`), with its paragraphs.
  const handedOn = (table: Table): ReadonlySet<TableCell> => {
    const grid = tableGrid(table, DOCUMENT_OP_TYPES.SET_TABLE_ROWS);
    if (grid.isErr()) return new Set();
    const kept = new Set<TableCell>();
    for (const [rowIndex, slots] of grid.value.rows.entries()) {
      if (removes(table.rows[rowIndex]?.structuralChange)) continue;
      for (const slot of slots) {
        if (slot.ownerRow === rowIndex) continue;
        const owner = table.rows[slot.ownerRow]?.cells[slot.ownerIndex];
        if (owner !== undefined) kept.add(owner);
      }
    }
    return kept;
  };
  // A nested table's paragraphs are in its own rows and in the outer cell
  // holding it: each paragraph is settled once, whichever removal reaches it.
  const settled = new Map<string, DocumentOp>();
  for (const table of tablesIn(document, op.story)) {
    const kept = handedOn(table);
    const mergeRows = reject ? undefined : selectedMergeMembers(table, ids);
    for (const row of table.rows)
      for (const cell of row.cells) {
        const goes =
          removes(cell.structuralChange) ||
          mergeClosesRow(row, mergeRows) ||
          (removes(row.structuralChange) && !kept.has(cell));
        if (!goes) continue;
        for (const { paragraph } of storyParagraphs({
          content: cell.content,
        })) {
          const { paraId } = paragraph;
          if (
            paragraph.pPrMark === undefined ||
            paraId === undefined ||
            settled.has(paraId)
          )
            continue;
          const review = reviewFieldsOf(paragraph);
          delete review.pPrMark;
          settled.set(paraId, {
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
            story: op.story,
            blockId: paraId,
            expected: reviewFieldsOf(paragraph),
            review,
          });
        }
      }
  }
  return [...settled.values()];
};

export const resolveTableRows = ({
  document,
  op,
  applyOps,
}: ResolveTableRowsOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  const ids = new Set(op.revisionIds);
  const reject = op.decision === REVISION_DECISIONS.REJECT;
  let current = document;
  const edits: DocumentEdit[] = [];

  const settled = marksInRemovedStructure(document, op);
  if (settled.length > 0) {
    const applied = applyOps(document, settled);
    if (applied.isErr()) return Result.err(applied.error);
    edits.push(applied.value);
    current = applied.value.document;
  }

  // Keep row insertion/deletion on its established primitive, which also
  // removes an emptied table. Table property and cell changes use setTable.
  for (const originalTable of tablesIn(document, op.story)) {
    // A merge already records the replacement cell topology. Applying its
    // row deletions separately would carry old merge owners into that topology.
    const mergeRows = reject
      ? undefined
      : selectedMergeMembers(originalTable, ids);
    if (
      !originalTable.rows.some(
        (row) =>
          row.structuralChange !== undefined &&
          ids.has(row.structuralChange.info.id) &&
          !mergeClosesRow(row, mergeRows),
      )
    )
      continue;
    const blockId = tableRowAnchor(originalTable.rows);
    if (blockId === undefined)
      return Result.err(
        refusal(op, "The table has no paragraph that can address its rows."),
      );
    const target = {
      type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
      story: op.story,
      blockId,
      expected: originalTable.rows,
      rows: originalTable.rows,
    } as const;
    const located = locateTableRow(current, target);
    if (located.isErr()) return Result.err(located.error);
    // A row the decision takes out hands the merges it starts to the next
    // row, as removing it directly does, so a merged cell keeps its content.
    const table = located.value.table;
    const rows = [...table.rows];
    let rowIndex = 0;
    while (rowIndex < rows.length) {
      const row = rows[rowIndex];
      const change = row?.structuralChange;
      if (
        row === undefined ||
        change === undefined ||
        !ids.has(change.info.id) ||
        mergeClosesRow(row, mergeRows)
      ) {
        rowIndex += 1;
        continue;
      }
      const added = change.type === "tableRowInsertion";
      if (added !== (op.decision === REVISION_DECISIONS.ACCEPT)) {
        const remaining = rowsWithoutRow({ ...table, rows }, rowIndex);
        rows.splice(0, rows.length, ...remaining);
        continue;
      }
      const next = { ...row };
      delete next.structuralChange;
      rows[rowIndex] = next;
      rowIndex += 1;
    }
    if (removalOrphansAMerge(table, rows))
      return Result.err(
        refusal(
          op,
          "A removed row hands the merges it starts to the next row with their content.",
        ),
      );
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
        refusal(
          op,
          "The table has no paragraph that can address its review records.",
        ),
      );
    const target = {
      type: DOCUMENT_OP_TYPES.SET_TABLE,
      story: op.story,
      blockId,
    } as const;
    const located = locateTableRow(current, target);
    if (located.isErr()) return Result.err(located.error);
    const table = located.value.table;
    // A decision that takes every cell away takes the table with them, as
    // removing its last row does; no grid is left to describe.
    if (removesEveryCell({ table, ids, reject })) {
      const applied = applyOps(current, [
        {
          type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
          story: op.story,
          blockId,
          expected: table.rows,
          rows: [],
        },
      ]);
      if (applied.isErr()) return Result.err(applied.error);
      edits.push(applied.value);
      current = applied.value.document;
      continue;
    }
    const grid = table.formatting?.gridChange;
    const selectedGridChange = grid !== undefined && ids.has(grid.id);
    const cellsResolved = resolveCellChanges({
      table,
      ids,
      reject,
      selectedGridChange,
    });
    if (cellsResolved.isErr())
      return Result.err(
        refusal(
          op,
          "The selected cell revisions do not describe complete grid columns.",
        ),
      );
    const propertiesResolved = propertyTable({
      table: cellsResolved.value.table,
      keepsGrid: cellsResolved.value.holdsRemovedColumns,
      ids,
      reject,
    });
    // Cells the decision removes take no bookmark beside them with them.
    const geometry = grid?.geometryChange;
    const nextTable = relocatePositionedMarkers(table, propertiesResolved, {
      mergedOwner: selectedGridChange ? pendingMergedOwner(table) : undefined,
      ...(!reject && selectedGridChange && geometry?.type === "mergeCells"
        ? { freshCellRevisions: new Set(geometry.revisionIds) }
        : {}),
    });
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
    ...(table.formatting?.gridChange === undefined
      ? []
      : [table.formatting.gridChange.id]),
    ...table.rows.flatMap((row) => [
      ...propertyIds(row.propertyChanges),
      ...propertyIds(row.tablePropertyExceptionChanges),
      ...(row.structuralChange === undefined
        ? []
        : [row.structuralChange.info.id]),
      ...row.cells.flatMap((cell) => [
        ...propertyIds(cell.propertyChanges),
        ...cellStructuralIds(cell),
      ]),
    ]),
  ]);

/**
 * The table once every pending record of its grid, its properties, its rows'
 * properties and its cells is accepted or rejected, as resolving them does;
 * the rows' own insertions and deletions stay pending. `unsupported` where
 * resolving them refuses.
 */
export const withTableRecordsResolved = (
  table: Table,
  reject: boolean,
): Result<Table, "unsupported"> => {
  const rowIds = new Set(
    table.rows.flatMap(({ structuralChange }) =>
      structuralChange === undefined ? [] : [structuralChange.info.id],
    ),
  );
  const ids = new Set(
    [...reachableOnTable(table)].filter((id) => !rowIds.has(id)),
  );
  if (ids.size === 0) return Result.ok(table);
  if (removesEveryCell({ table, ids, reject }))
    return Result.err("unsupported");
  const cells = resolveCellChanges({
    table,
    ids,
    reject,
    selectedGridChange:
      table.formatting?.gridChange !== undefined &&
      ids.has(table.formatting.gridChange.id),
  });
  if (cells.isErr()) return Result.err(cells.error);
  return Result.ok(
    propertyTable({
      table: cells.value.table,
      keepsGrid: cells.value.holdsRemovedColumns,
      ids,
      reject,
    }),
  );
};
