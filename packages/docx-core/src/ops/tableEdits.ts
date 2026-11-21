/** Semantic grid/cell/property edits; one exact, stale-checked table inverse. */
import { panic, Result } from "better-result";

import {
  type Document,
  MAX_REVISION_ID,
  type Table,
  type TableCell,
  type TableCellFormatting,
  type TableFormatting,
  type TableMeasurement,
  type TableRow,
  type TableRowFormatting,
} from "../model/document";
import {
  storyBody,
  storyParagraphs,
  updateBlockList,
  withBodyContent,
} from "./blocks";
import { validateOpsDocument } from "./contract";
import type { DocumentEdit } from "./edits";
import { equalForStaleness, structurallyEqual } from "./equality";
import { idKey, isParaId, packageParagraphIds, paragraphIdsIn } from "./ids";
import { applyFormattingPatch } from "./patch";
import { relocatePositionedMarkers } from "./positionedMarkers";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { stampInfo, WRAP_KINDS } from "./review";
import { replaceStoryBody } from "./stories";
import {
  reachesPastRemovedCell,
  strandsABookmark,
} from "./tableEditDependencies";
import { editRemovesRow, type TableGrid, tableGrid } from "./tableGrid";
import {
  locateTableRow,
  tableRowAnchor,
  type TableRowLocation,
} from "./tableLocation";
import { applyTableOp } from "./tables";
import { trackTableRows } from "./tableTracking";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOpType,
  type RevisionStamp,
  type SetTableOp,
  type TableEditOp,
  type TableIntentOperation,
} from "./types";

const CELL_VERTICAL_MERGES = {
  RESTART: "restart",
  CONTINUE: "continue",
} as const;

type Op = TableEditOp | SetTableOp;
type RefusalOptions = {
  op: Pick<Op, "type">;
  reason: DocumentOpRefusal["reason"];
  message: string;
};
const refuse = ({ op, reason, message }: RefusalOptions) =>
  Result.err(new DocumentOpRefusal({ opType: op.type, reason, message }));
const mismatch = (op: Pick<Op, "type">, message: string) =>
  refuse({
    op,
    reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    message,
  });

const patchChanges = (base: object | undefined, patch: object): boolean =>
  Object.keys(patch).some((key) => {
    const value: unknown = Reflect.get(patch, key);
    if (value === undefined) return false;
    const previous: unknown =
      base === undefined ? undefined : Reflect.get(base, key);
    return value === null
      ? previous !== undefined
      : !structurallyEqual(previous, value);
  });

/** Source tokens describe the old property set and must not survive an authored edit. */
const authored = <T extends { sourceXml?: string }>(formatting: T): T => {
  const next = { ...formatting };
  delete next.sourceXml;
  return next;
};

type FreshCellsOptions = {
  document: Document;
  op: Op;
  ids: readonly string[];
  count: number;
};
const freshCells = ({
  document,
  op,
  ids,
  count,
}: FreshCellsOptions): Result<TableCell[], DocumentOpRefusal> => {
  if (ids.length !== count)
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
      message: `This edit requires ${count} fresh paragraph ids.`,
    });
  if (ids.some((id) => !isParaId(id)))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      message: "New table paragraphs need usable paragraph ids.",
    });
  const occupied = new Set(packageParagraphIds(document.package).map(idKey));
  for (const id of ids) {
    if (occupied.has(idKey(id)))
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
        message: "A new table paragraph id is already in use.",
      });
    occupied.add(idKey(id));
  }
  return Result.ok(
    ids.map((paraId) => ({
      type: "tableCell",
      content: [{ type: "paragraph", paraId, content: [] }],
    })),
  );
};

const PROPERTY_FIELDS = {
  cell: [
    "width",
    "borders",
    "margins",
    "shading",
    "verticalAlign",
    "textDirection",
    "fitText",
    "noWrap",
    "hideMark",
    "conditionalFormat",
  ] as const satisfies readonly (keyof TableCellFormatting)[],
  row: [
    "widthBefore",
    "widthAfter",
    "height",
    "heightRule",
    "header",
    "cantSplit",
    "justification",
    "hidden",
    "conditionalFormat",
  ] as const satisfies readonly (keyof TableRowFormatting)[],
  table: [
    "width",
    "justification",
    "cellSpacing",
    "indent",
    "borders",
    "cellMargins",
    "layout",
    "styleId",
    "look",
    "shading",
    "overlap",
    "floating",
    "bidi",
    "caption",
    "description",
    "rowBandSize",
    "columnBandSize",
  ] as const satisfies readonly (keyof TableFormatting)[],
};

type PropertyEdit = Extract<
  TableEditOp,
  {
    type:
      | typeof DOCUMENT_OP_TYPES.SET_CELL_PROPS
      | typeof DOCUMENT_OP_TYPES.SET_ROW_PROPS
      | typeof DOCUMENT_OP_TYPES.SET_TABLE_PROPS;
  }
>;

const isPropertyEdit = (op: Op | TableIntentOperation): op is PropertyEdit => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
    case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
    case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
      return true;
    default:
      return false;
  }
};

/** Remove only the properties these operations can change; every other field remains evidence. */
const structuralFormatting = (
  formatting: object | undefined,
  fields: readonly string[],
) => {
  const entries = Object.entries(formatting ?? {}).filter(
    ([key]) => key !== "sourceXml" && !fields.includes(key),
  );
  return entries.length === 0 ? undefined : Object.fromEntries(entries);
};

type ReviewStructureOptions = {
  formatting: object | undefined;
  fields: readonly string[];
  changes:
    | readonly {
        type: string;
        info: unknown;
        previousFormatting?: object;
        currentFormatting?: object;
      }[]
    | undefined;
};
/** Property-only history may appear/disappear; recorded structural history may not change. */
const reviewStructure = ({
  formatting,
  fields,
  changes,
}: ReviewStructureOptions) => {
  const current = structuralFormatting(formatting, fields);
  return (changes ?? []).flatMap(
    ({ info: _info, previousFormatting, currentFormatting, ...review }) => {
      const previous = structuralFormatting(previousFormatting, fields);
      const recordedCurrent =
        currentFormatting === undefined
          ? current
          : structuralFormatting(currentFormatting, fields);
      if (
        Object.keys(review).length === 1 &&
        structurallyEqual(previous, current) &&
        structurallyEqual(recordedCurrent, current)
      )
        return [];
      return [
        Object.assign({}, review, {
          previousFormatting: previous,
          currentFormatting: recordedCurrent,
        }),
      ];
    },
  );
};

const propertyEditStructure = ({
  formatting: tableFormatting,
  propertyChanges: tableChanges,
  rows,
  ...table
}: Table) =>
  Object.assign({}, table, {
    formatting: structuralFormatting(tableFormatting, PROPERTY_FIELDS.table),
    propertyChanges: reviewStructure({
      formatting: tableFormatting,
      fields: PROPERTY_FIELDS.table,
      changes: tableChanges,
    }),
    rows: rows.map(
      ({
        formatting: rowFormatting,
        propertyChanges: rowChanges,
        cells,
        ...row
      }) =>
        Object.assign({}, row, {
          formatting: structuralFormatting(rowFormatting, PROPERTY_FIELDS.row),
          propertyChanges: reviewStructure({
            formatting: rowFormatting,
            fields: PROPERTY_FIELDS.row,
            changes: rowChanges,
          }),
          cells: cells.map(
            ({
              formatting: cellFormatting,
              propertyChanges: cellChanges,
              ...cell
            }) =>
              Object.assign({}, cell, {
                formatting: structuralFormatting(
                  cellFormatting,
                  PROPERTY_FIELDS.cell,
                ),
                propertyChanges: reviewStructure({
                  formatting: cellFormatting,
                  fields: PROPERTY_FIELDS.cell,
                  changes: cellChanges,
                }),
              }),
          ),
        }),
    ),
  });

type CommitTableOptions = {
  document: Document;
  op: Op;
  location: TableRowLocation;
  table: Table;
};
const commitTable = ({
  document,
  op,
  location,
  table: edited,
}: CommitTableOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  // An inverse restores the table exactly; an edit moves the bookmark markers
  // beside rows and cells with the siblings it adds or removes.
  const table =
    op.type === DOCUMENT_OP_TYPES.SET_TABLE
      ? edited
      : relocatePositionedMarkers(location.table, edited);
  const preservesStructure =
    (isPropertyEdit(op) || op.type === DOCUMENT_OP_TYPES.SET_TABLE) &&
    structurallyEqual(
      propertyEditStructure(location.table),
      propertyEditStructure(table),
    );
  if (isPropertyEdit(op)) {
    if (!preservesStructure)
      return mismatch(
        op,
        "A property edit must preserve table structure, content and identities.",
      );
  } else if (op.type !== DOCUMENT_OP_TYPES.SET_TABLE || !preservesStructure) {
    const checked = tableGrid(table, op.type);
    if (checked.isErr()) return Result.err(checked.error);
  }
  if (
    table.rows.some((row) =>
      row.cells.some(
        (cell) =>
          !storyParagraphs({ content: cell.content }).some(({ list }) =>
            list.every((step) => step.kind !== "tableCell"),
          ),
      ),
    )
  ) {
    return mismatch(
      op,
      "Each cell must retain a paragraph outside its nested tables.",
    );
  }
  const anchor = tableRowAnchor(table.rows);
  if (anchor === undefined)
    return mismatch(op, "The resulting table has no paragraph anchor.");
  const before = location.table;
  const beforeIds = paragraphIdsIn(before);
  const afterIds = paragraphIdsIn(table);
  const old = new Set(beforeIds.map(idKey));
  const nextIds = new Set(afterIds.map(idKey));
  if (afterIds.some((id) => !old.has(idKey(id)) && !isParaId(id)))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      message: "An added table paragraph has an invalid id.",
    });
  const body = storyBody(document, op.story);
  const content = updateBlockList(body.content, location.list, (blocks) => {
    const next = [...blocks];
    next[location.index] = table;
    return next;
  });
  const next = replaceStoryBody({
    document,
    story: op.story,
    body: withBodyContent(body, content),
  });
  const valid = validateOpsDocument(next);
  if (valid.isErr())
    return refuse({
      op,
      reason: valid.error.reason,
      message: valid.error.message,
    });
  return Result.ok({
    document: next,
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.SET_TABLE,
        story: op.story,
        blockId: anchor,
        expected: table,
        table: before,
      },
    ],
    touched: {
      modified: beforeIds.filter((id) => nextIds.has(idKey(id))),
      inserted: afterIds.filter((id) => !old.has(idKey(id))),
      removed: beforeIds.filter((id) => !nextIds.has(idKey(id))),
    },
  });
};

const withSpan = (cell: TableCell, span: number): TableCell => {
  const formatting = authored({ ...cell.formatting });
  if (span === 1) delete formatting.gridSpan;
  else formatting.gridSpan = span;
  return { ...cell, formatting };
};
const dividedWidth = (
  width: TableMeasurement | undefined,
  parts: number,
): TableMeasurement[] | undefined => {
  if (width === undefined) return undefined;
  if (width.type !== "dxa" && width.type !== "pct")
    return Array.from({ length: parts }, () => width);
  const base = Math.floor(width.value / parts);
  const remainder = width.value - base * parts;
  return Array.from({ length: parts }, (_, index) => ({
    type: width.type,
    value: base + (index < remainder ? 1 : 0),
  }));
};
const mergedWidth = (
  cells: readonly TableCell[],
): Result<TableMeasurement | undefined, "invalid"> => {
  const first = cells.at(0)?.formatting?.width;
  if (first === undefined || (first.type !== "dxa" && first.type !== "pct"))
    return Result.ok(first);
  if (!Number.isSafeInteger(first.value) || first.value < 0)
    return Result.err("invalid");
  let value = first.value;
  for (const cell of cells.slice(1)) {
    const width = cell.formatting?.width;
    // Unknown or incompatible widths do not add to a unit-bearing leading width.
    if (width === undefined || width.type !== first.type)
      return Result.ok(first);
    if (!Number.isSafeInteger(width.value) || width.value < 0)
      return Result.err("invalid");
    value += width.value;
    if (!Number.isSafeInteger(value)) return Result.err("invalid");
  }
  return Result.ok({ type: first.type, value });
};
const withOmission = (
  row: TableRow,
  side: "gridBefore" | "gridAfter",
  value: number,
): TableRow => {
  const formatting = authored({ ...row.formatting, [side]: value });
  if (value === 0) {
    if (side === "gridBefore") delete formatting.gridBefore;
    else delete formatting.gridAfter;
  }
  return { ...row, formatting };
};

const hasStructuralMarkup = (table: Table) =>
  table.preserved !== undefined ||
  table.bookmarks !== undefined ||
  table.carrierStack !== undefined ||
  table.rows.some(
    (row) =>
      row.preserved !== undefined ||
      row.carrierStack !== undefined ||
      row.contentControls !== undefined,
  );
/** A cell holding one paragraph with nothing in it. */
const isEmptyCell = (cell: TableCell) =>
  cell.content.length === 1 &&
  cell.content[0]?.type === "paragraph" &&
  cell.content[0].content.length === 0;
const hasReview = (table: Table) =>
  table.formatting?.gridChange !== undefined ||
  table.propertyChanges !== undefined ||
  table.rows.some(
    (row) =>
      row.structuralChange !== undefined ||
      row.propertyChanges !== undefined ||
      row.cells.some(
        (cell) =>
          cell.structuralChange !== undefined ||
          cell.propertyChanges !== undefined,
      ),
  );

type HasReviewBlockingMergeOptions = {
  table: Table;
  /** The first row the merge changes. */
  top: number;
  /** The row after the last one the merge changes. */
  bottom: number;
};

/**
 * A review a cell merge cannot join: any but a row's own suggested insertion
 * or deletion outside the rows the merge changes, which resolves apart from
 * the merge, as the merge's rows resolve apart from it.
 */
const hasReviewBlockingMerge = ({
  table,
  top,
  bottom,
}: HasReviewBlockingMergeOptions) =>
  table.formatting?.gridChange !== undefined ||
  table.propertyChanges !== undefined ||
  table.rows.some(
    (row, index) =>
      (row.structuralChange !== undefined && index >= top && index < bottom) ||
      row.propertyChanges !== undefined ||
      row.cells.some(
        (cell) =>
          cell.structuralChange !== undefined ||
          cell.propertyChanges !== undefined,
      ),
  );

/**
 * A review a column edit cannot join. A row's own suggested insertion or
 * deletion keeps its mark while the column gains or loses that row's cell;
 * another suggested column deletion may join pending cell deletions.
 */
const hasReviewBlockingColumns = (table: Table, suggestedDeletion: boolean) =>
  table.formatting?.gridChange !== undefined ||
  table.propertyChanges !== undefined ||
  table.rows.some(
    (row) =>
      row.propertyChanges !== undefined ||
      row.cells.some(
        (cell) =>
          (cell.structuralChange !== undefined &&
            !(
              suggestedDeletion &&
              cell.structuralChange.type === "tableCellDeletion"
            )) ||
          cell.propertyChanges !== undefined,
      ),
  );

/** Multiple physical review records always get distinct, explicitly supplied ids. */
const stampsFor = (
  op: TableEditOp,
  count: number,
): Result<RevisionStamp[], DocumentOpRefusal> => {
  const revision = op.revision;
  if (revision === undefined) return Result.ok([]);
  const ids = [revision.id, ...(op.newIds?.revision ?? [])];
  if (new Set(ids).size !== ids.length)
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      message: "Revision ids must be distinct within the operation.",
    });
  if (
    ids
      .slice(0, count)
      .some((id) => !Number.isInteger(id) || id < 0 || id > MAX_REVISION_ID)
  )
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
      message: "Revision ids must be nonnegative 31-bit integers.",
    });
  if (ids.length < count)
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
      message: `This table change needs ${count} distinct revision ids.`,
    });
  return Result.ok(
    ids.slice(0, count).map((id) => Object.assign({}, revision, { id })),
  );
};

/**
 * Where a row's cells stand: after what it omits before them, up to where
 * they end. A row may stop short of the grid by more than its `w:gridAfter`
 * states; the columns after its cells are omitted either way.
 */
const rowExtent = (table: Table, grid: TableGrid, row: number) => {
  const before = table.rows[row]?.formatting?.gridBefore ?? 0;
  return { before, end: grid.rows[row]?.at(-1)?.end ?? before };
};

type ColumnOptions = {
  document: Document;
  op: Extract<TableEditOp, { type: "insertColumn" | "deleteColumn" }>;
  table: Table;
  grid: TableGrid;
};
const columnEdit = ({
  document,
  op,
  table,
  grid,
}: ColumnOptions): Result<Table, DocumentOpRefusal> => {
  const insert = op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN;
  if (
    !Number.isSafeInteger(op.column) ||
    op.column < 0 ||
    op.column >= grid.width + (insert ? 1 : 0)
  )
    return mismatch(op, "Column is outside the logical grid.");
  if (!insert && grid.width === 1)
    return mismatch(
      op,
      "Deleting the final column requires deleting the table.",
    );
  if (insert && (!Number.isSafeInteger(op.width) || op.width <= 0))
    return mismatch(op, "A column width must be a positive twip count.");
  if (hasStructuralMarkup(table))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message:
        "Indexed table markup must be relocated before changing columns.",
    });
  // Another suggested column deletion may join pending ones: each keeps its
  // own cells and grid entry until it is resolved.
  const suggestedDeletion = !insert && op.revision !== undefined;
  if (hasReviewBlockingColumns(table, suggestedDeletion))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      message: "Resolve table reviews before changing its columns.",
    });
  // The cells the column removes, now or when its suggestion is accepted,
  // take their bookmarks, content controls and own review records with them.
  const removedCells = insert
    ? []
    : grid.rows.flatMap((entries) =>
        entries
          .filter(
            (entry) =>
              entry.start <= op.column &&
              op.column < entry.end &&
              entry.end - entry.start === 1,
          )
          .map((entry) => entry.cell),
      );
  if (
    removedCells.some(reachesPastRemovedCell) ||
    strandsABookmark(removedCells)
  )
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.DEPENDENT_RECORDS,
      message:
        "Resolve comments, moves and bookmarks reaching past the column before removing it.",
    });
  const boundaryRows = grid.rows.filter((cells, row) => {
    const { before, end } = rowExtent(table, grid, row);
    return (
      op.column >= before &&
      op.column <= end &&
      !cells.some((cell) => cell.start < op.column && op.column < cell.end)
    );
  }).length;
  const empty: TableCell[] = [];
  const fresh =
    op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN
      ? freshCells({ document, op, ids: op.newBlockIds, count: boundaryRows })
      : Result.ok(empty);
  if (fresh.isErr()) return Result.err(fresh.error);
  // A suggested deletion keeps the grid while it is pending: the deleted
  // cells stay in their columns, and accepting removes them with their grid
  // entries.
  const trackedDeletion = !insert && op.revision !== undefined;
  // The rows the edit leaves as they are: for a suggested deletion, those
  // with no cell in the column, which keep it for themselves when the
  // deletion is accepted; for another edit, those stopping short of the
  // column by more than they state.
  const leavesRow = (rowIndex: number) => {
    const { before, end } = rowExtent(table, grid, rowIndex);
    if (trackedDeletion) return op.column < before || op.column >= end;
    const stated = table.rows[rowIndex]?.formatting?.gridAfter ?? 0;
    return insert ? op.column > end + stated : op.column >= end + stated;
  };
  // Each row the edit changes takes a stamp; an edit besides a suggested
  // deletion stamps the grid and table properties after them.
  const stampedRows = new Map<number, number>();
  for (const rowIndex of table.rows.keys())
    if (!leavesRow(rowIndex)) stampedRows.set(rowIndex, stampedRows.size);
  if (trackedDeletion && stampedRows.size === 0)
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message:
        "A suggested column deletion needs a cell in the column to mark.",
    });
  const stamps = stampsFor(
    op,
    trackedDeletion ? stampedRows.size : stampedRows.size + 2,
  );
  if (stamps.isErr()) return Result.err(stamps.error);
  const rowStamp = (rowIndex: number) =>
    stamps.value[stampedRows.get(rowIndex) ?? -1];
  let freshIndex = 0;
  const rows: TableRow[] = [];
  for (const [rowIndex, row] of table.rows.entries()) {
    const entries = grid.rows[rowIndex];
    if (entries === undefined) return mismatch(op, "The grid row is missing.");
    if (leavesRow(rowIndex)) {
      rows.push(row);
      continue;
    }
    const { before, end } = rowExtent(table, grid, rowIndex);
    if (op.column < before) {
      const next = withOmission(row, "gridBefore", before + (insert ? 1 : -1));
      const stamp = rowStamp(rowIndex);
      if (stamp !== undefined)
        next.propertyChanges = [
          {
            type: "tableRowPropertyChange",
            info: stampInfo(stamp),
            ...(row.formatting === undefined
              ? {}
              : { previousFormatting: row.formatting }),
          },
        ];
      rows.push(next);
      continue;
    }
    if (op.column >= end && (!insert || op.column > end)) {
      const next = withOmission(
        row,
        "gridAfter",
        (row.formatting?.gridAfter ?? 0) + (insert ? 1 : -1),
      );
      const stamp = rowStamp(rowIndex);
      if (stamp !== undefined)
        next.propertyChanges = [
          {
            type: "tableRowPropertyChange",
            info: stampInfo(stamp),
            ...(row.formatting === undefined
              ? {}
              : { previousFormatting: row.formatting }),
          },
        ];
      rows.push(next);
      continue;
    }
    const cells: TableCell[] = [];
    let inserted = false;
    for (const entry of entries) {
      if (insert && entry.start === op.column) {
        const cell = fresh.value[freshIndex++];
        if (cell === undefined) return mismatch(op, "A fresh cell is missing.");
        const stamp = rowStamp(rowIndex);
        cells.push(
          stamp === undefined
            ? cell
            : {
                ...cell,
                structuralChange: {
                  type: "tableCellInsertion",
                  info: stampInfo(stamp),
                },
              },
        );
        inserted = true;
      }
      const hit = entry.start <= op.column && op.column < entry.end;
      if (insert && entry.start < op.column && hit) {
        const next = withSpan(entry.cell, entry.end - entry.start + 1);
        const stamp = rowStamp(rowIndex);
        if (stamp !== undefined)
          next.propertyChanges = [
            {
              type: "tableCellPropertyChange",
              info: stampInfo(stamp),
              ...(entry.cell.formatting === undefined
                ? {}
                : { previousFormatting: entry.cell.formatting }),
            },
          ];
        cells.push(next);
        inserted = true;
        continue;
      }
      if (!insert && hit) {
        if (
          trackedDeletion &&
          entry.cell.structuralChange?.type === "tableCellDeletion"
        )
          return refuse({
            op,
            reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
            message: "The column's cells are already suggested for deletion.",
          });
        if (trackedDeletion && entry.end - entry.start > 1)
          return refuse({
            op,
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message:
              "A suggested column deletion cannot narrow a cell that spans other columns.",
          });
        if (entry.end - entry.start > 1) {
          const next = withSpan(entry.cell, entry.end - entry.start - 1);
          const stamp = rowStamp(rowIndex);
          if (stamp !== undefined)
            next.propertyChanges = [
              {
                type: "tableCellPropertyChange",
                info: stampInfo(stamp),
                ...(entry.cell.formatting === undefined
                  ? {}
                  : { previousFormatting: entry.cell.formatting }),
              },
            ];
          cells.push(next);
          continue;
        }
        const stamp = rowStamp(rowIndex);
        if (stamp !== undefined)
          cells.push({
            ...entry.cell,
            structuralChange: {
              type: "tableCellDeletion",
              info: stampInfo(stamp),
            },
          });
        continue;
      }
      cells.push(entry.cell);
    }
    if (insert && !inserted) {
      const cell = fresh.value[freshIndex++];
      if (cell === undefined) return mismatch(op, "A fresh cell is missing.");
      const stamp = rowStamp(rowIndex);
      cells.push(
        stamp === undefined
          ? cell
          : {
              ...cell,
              structuralChange: {
                type: "tableCellInsertion",
                info: stampInfo(stamp),
              },
            },
      );
    }
    // A row the deletion takes its only cells, or the only cells it
    // starts, from goes with them; a suggested deletion leaves it to go
    // when the suggestion is accepted.
    if (editRemovesRow({ before: row.cells, after: cells })) continue;
    rows.push({ ...row, cells });
  }
  if (trackedDeletion) return Result.ok({ ...table, rows });
  const columnWidths =
    table.columnWidths === undefined ? undefined : [...table.columnWidths];
  if (columnWidths !== undefined) {
    if (op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN)
      columnWidths.splice(op.column, 0, op.width);
    else columnWidths.splice(op.column, 1);
  }
  const next = { ...table, rows };
  if (columnWidths !== undefined) next.columnWidths = columnWidths;
  const gridStamp = stamps.value[stampedRows.size];
  const propertyStamp = stamps.value[stampedRows.size + 1];
  if (
    op.revision !== undefined &&
    gridStamp !== undefined &&
    propertyStamp !== undefined
  ) {
    next.formatting = {
      ...table.formatting,
      // A grid the file states no widths for is recorded as its columns,
      // each without a measure, as the file states them.
      gridChange: {
        id: gridStamp.id,
        columnWidths:
          table.columnWidths === undefined
            ? Array.from({ length: grid.width }, () => undefined)
            : [...table.columnWidths],
      },
    };
    next.propertyChanges = [
      {
        type: "tablePropertyChange",
        info: stampInfo(propertyStamp),
        ...(table.formatting === undefined
          ? {}
          : { previousFormatting: table.formatting }),
      },
    ];
  }
  return Result.ok(next);
};

type MergeSelectionOptions = {
  table: Table;
  grid: TableGrid;
  op: Extract<TableIntentOperation, { type: "mergeCells" }>;
};
const mergeSelection = ({ table, grid, op }: MergeSelectionOptions) => {
  if (
    ![op.top, op.bottom, op.left, op.right].every(Number.isSafeInteger) ||
    op.top < 0 ||
    op.bottom > table.rows.length ||
    op.left < 0 ||
    op.right > grid.width ||
    op.top >= op.bottom ||
    op.left >= op.right
  )
    return mismatch(op, "The merge rectangle is outside the table.");
  const selected = grid.rows
    .slice(op.top, op.bottom)
    .map((row) =>
      row.filter((cell) => cell.start < op.right && cell.end > op.left),
    );
  for (const row of selected) {
    if (
      row.at(0)?.start !== op.left ||
      row.at(-1)?.end !== op.right ||
      row.some(
        (cell) =>
          cell.start < op.left || cell.end > op.right || cell.ownerRow < op.top,
      )
    )
      return mismatch(op, "The merge rectangle cuts a span or omitted slot.");
  }
  if (
    grid.rows
      .slice(op.bottom)
      .some((row) =>
        row.some(
          (cell) =>
            cell.ownerRow >= op.top &&
            cell.ownerRow < op.bottom &&
            cell.start < op.right &&
            cell.end > op.left,
        ),
      )
  )
    return mismatch(op, "The merge rectangle cuts a vertical span.");
  // A pure vertical merge of empty continuations changes only merge state. Its
  // cellMerge records can reject exactly without inventing a content-move history.
  const verticalOnly =
    op.bottom - op.top > 1 &&
    selected.every((row) => row.length === 1) &&
    selected
      .slice(1)
      .every((row) =>
        row.every(({ cell }) =>
          cell.content.every(
            (block) => block.type === "paragraph" && block.content.length === 0,
          ),
        ),
      );
  return Result.ok({ selected, verticalOnly });
};

type ClosingRowsOptions = {
  /** The table's rows before the merge. */
  before: readonly TableRow[];
  /** The rows the merge leaves, one for each of `before`. */
  after: readonly TableRow[];
};

/**
 * The rows the merge leaves nothing they start, only merges from above
 * (`editRemovesRow`). A reader cannot hold such a row, so the merge removes
 * it and closes over one row fewer.
 */
const closingRows = ({ before, after }: ClosingRowsOptions): Set<number> =>
  new Set(
    after.flatMap((row, index) =>
      editRemovesRow({ before: before[index]?.cells ?? [], after: row.cells })
        ? [index]
        : [],
    ),
  );

type TrackedClosingRowsOptions = {
  document: Document;
  op: Extract<TableEditOp, { type: "mergeCells" }>;
  table: Table;
  /** The merged rows, one for each of the table's. */
  merged: readonly TableRow[];
  closing: ReadonlySet<number>;
  /** The revision ids the merge's own cell records take. */
  used: number;
};

/**
 * A suggested merge's rows, with the deletion of each row it closes
 * recorded beside its cell's merge record: accepting removes the row, as the
 * direct merge does, and rejecting keeps it as it was.
 */
const trackedClosingRows = ({
  document,
  op,
  table,
  merged,
  closing,
  used,
}: TrackedClosingRowsOptions): Result<TableRow[], DocumentOpRefusal> => {
  const stamps = stampsFor(op, used + 1);
  if (stamps.isErr()) return Result.err(stamps.error);
  const revision =
    stamps.value.at(-1) ?? panic("A suggested merge lost its row stamp.");
  const indexes = [...closing];
  const tracked = trackTableRows({
    document,
    op,
    rows: indexes.map(
      (index) => table.rows[index] ?? panic("A closed row is missing."),
    ),
    revision,
    newIds: {
      ...op.newIds,
      revision: (op.newIds?.revision ?? []).slice(used),
    },
    kind: WRAP_KINDS.DELETION,
  });
  if (tracked.isErr()) return Result.err(tracked.error);
  const rows = [...merged];
  for (const [position, index] of indexes.entries()) {
    const deleted =
      tracked.value[position] ?? panic("A suggested merge lost a row.");
    const row = merged[index] ?? panic("A merged row is missing.");
    const cells: TableCell[] = [];
    for (const [cellIndex, cell] of deleted.cells.entries())
      cells.push(
        Object.assign(
          {},
          row.cells[cellIndex] ?? panic("A merged cell is missing."),
          { content: cell.content },
        ),
      );
    rows[index] = { ...deleted, cells };
  }
  return Result.ok(rows);
};

type MergeOptions = {
  document: Document;
  op: Extract<TableEditOp, { type: "mergeCells" }>;
  table: Table;
  grid: TableGrid;
};
const merge = ({
  document,
  op,
  table,
  grid,
}: MergeOptions): Result<Table, DocumentOpRefusal> => {
  if (hasStructuralMarkup(table))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "Indexed table markup must be relocated before changing cells.",
    });
  if (hasReviewBlockingMerge({ table, top: op.top, bottom: op.bottom }))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      message: "Resolve table reviews before changing cells.",
    });
  const selection = mergeSelection({ table, grid, op });
  if (selection.isErr()) return Result.err(selection.error);
  const { selected, verticalOnly } = selection.value;
  if (verticalOnly) {
    const fresh = freshCells({ document, op, ids: op.newBlockIds, count: 0 });
    if (fresh.isErr()) return Result.err(fresh.error);
    const stamps = stampsFor(op, selected.length * 2);
    if (stamps.isErr()) return Result.err(stamps.error);
    const rows = table.rows.map((row, rowIndex) => {
      const entry = selected[rowIndex - op.top]?.at(0);
      if (entry === undefined) return row;
      const next: TableCell = {
        ...entry.cell,
        formatting: authored({
          ...entry.cell.formatting,
          vMerge:
            rowIndex === op.top
              ? CELL_VERTICAL_MERGES.RESTART
              : CELL_VERTICAL_MERGES.CONTINUE,
        }),
      };
      const stamp = stamps.value[(rowIndex - op.top) * 2];
      const propertyStamp = stamps.value[(rowIndex - op.top) * 2 + 1];
      if (stamp !== undefined)
        next.structuralChange = {
          type: "tableCellMerge",
          info: stampInfo(stamp),
          verticalMerge: rowIndex === op.top ? "rest" : "continue",
          ...(entry.cell.formatting?.vMerge === undefined
            ? {}
            : {
                verticalMergeOriginal:
                  entry.cell.formatting.vMerge === "restart"
                    ? "rest"
                    : "continue",
              }),
        };
      if (propertyStamp !== undefined)
        next.propertyChanges = [
          {
            type: "tableCellPropertyChange",
            info: stampInfo(propertyStamp),
            ...(entry.cell.formatting === undefined
              ? {}
              : { previousFormatting: entry.cell.formatting }),
          },
        ];
      const cells = [...row.cells];
      cells[entry.index] = next;
      return { ...row, cells };
    });
    const closing = closingRows({ before: table.rows, after: rows });
    if (closing.size === 0) return Result.ok({ ...table, rows });
    if (op.revision === undefined)
      return Result.ok({
        ...table,
        rows: rows.filter((_row, index) => !closing.has(index)),
      });
    return trackedClosingRows({
      document,
      op,
      table,
      merged: rows,
      closing,
      used: selected.length * 2,
    }).map((tracked) => Object.assign({}, table, { rows: tracked }));
  }
  if (op.revision !== undefined)
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message:
        "Cell merge markup cannot record horizontal content relocation; tracked moves are required.",
    });
  const fresh = freshCells({
    document,
    op,
    ids: op.newBlockIds,
    count: op.bottom - op.top - 1,
  });
  if (fresh.isErr()) return Result.err(fresh.error);
  const first = selected.at(0)?.at(0);
  if (first === undefined) return mismatch(op, "No cells are selected.");
  // A cell holding one empty paragraph gives the merged cell nothing; the
  // first cell's content stays when it holds any, or when no cell does.
  const merged = selected.flatMap((row) => row.map(({ cell }) => cell));
  const holding = merged.filter((cell) => !isEmptyCell(cell));
  const content = (holding.length === 0 ? [first.cell] : holding).flatMap(
    (cell) => cell.content,
  );
  const mergedCellWidth = mergedWidth(
    selected.at(0)?.map(({ cell }) => cell) ?? [],
  );
  if (mergedCellWidth.isErr())
    return mismatch(
      op,
      "Selected cell widths must be nonnegative safe integer units.",
    );
  const width = mergedCellWidth.value;
  const rows = table.rows.map((row, index) => {
    const entries = selected[index - op.top];
    if (entries === undefined) return row;
    const leading = entries.at(0);
    if (leading === undefined) return row;
    const base =
      index === op.top
        ? { ...first.cell, content }
        : fresh.value[index - op.top - 1];
    if (base === undefined) return row;
    const cell = withSpan(base, op.right - op.left);
    const formatting = authored({ ...cell.formatting });
    if (index === op.top) {
      if (width === undefined) delete formatting.width;
      else formatting.width = width;
    }
    if (op.bottom - op.top > 1)
      formatting.vMerge = index === op.top ? "restart" : "continue";
    else delete formatting.vMerge;
    const cells = [...row.cells];
    cells.splice(leading.index, entries.length, { ...cell, formatting });
    return { ...row, cells };
  });
  const closing = closingRows({ before: table.rows, after: rows });
  return Result.ok({
    ...table,
    rows: rows.filter((_row, index) => !closing.has(index)),
  });
};

type SplitOptions = {
  document: Document;
  op: Extract<TableEditOp, { type: "splitCell" }>;
  table: Table;
  grid: TableGrid;
  location: TableRowLocation;
};
const split = ({
  document,
  op,
  table,
  grid,
  location,
}: SplitOptions): Result<Table, DocumentOpRefusal> => {
  const target = grid.rows[location.rowIndex]?.find(
    (entry) => entry.index === location.cellIndex,
  );
  if (!target) return mismatch(op, "The addressed cell is absent.");
  if (target.row !== target.ownerRow || target.index !== target.ownerIndex)
    return mismatch(
      op,
      "A vertical merge continuation must be addressed through its owner cell.",
    );
  if (hasStructuralMarkup(table))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "Indexed table markup must be relocated before changing cells.",
    });
  if (hasReview(table))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      message: "Resolve table reviews before changing cells.",
    });
  const group = grid.rows.flatMap((row) =>
    row.filter(
      (cell) =>
        cell.ownerRow === target.ownerRow &&
        cell.ownerIndex === target.ownerIndex,
    ),
  );
  // A vertical merge splits into its rows' own cells, each of which keeps
  // its content and identity, as a span splits into its columns.
  const span = target.end - target.start;
  if (span === 1 && group.length === 1)
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE,
      message: "A one-row, one-column cell has nothing to split.",
    });
  if (
    group.some(({ cell }) => {
      const width = cell.formatting?.width;
      return (
        width !== undefined &&
        (width.type === "dxa" || width.type === "pct") &&
        (!Number.isSafeInteger(width.value) || width.value < 0)
      );
    })
  )
    return mismatch(
      op,
      "A split cell width must be a nonnegative safe integer unit.",
    );
  const fresh = freshCells({
    document,
    op,
    ids: op.newBlockIds,
    count: group.length * (span - 1),
  });
  if (fresh.isErr()) return Result.err(fresh.error);
  // A vertical merge records each row's leaving it as a cell merge change,
  // which rejecting reverts to the merge state the cell had.
  const vertical = group.length > 1;
  const stamps = stampsFor(op, group.length * (span + (vertical ? 1 : 0)));
  if (stamps.isErr()) return Result.err(stamps.error);
  let stampIndex = 0;
  let freshIndex = 0;
  const rows = table.rows.map((row, index) => {
    const entry = group.find((cell) => cell.row === index);
    if (!entry) return row;
    const formatting = authored({ ...entry.cell.formatting });
    delete formatting.gridSpan;
    delete formatting.vMerge;
    const widths = dividedWidth(entry.cell.formatting?.width, span);
    const formattingForColumn = (column: number) => {
      const next = { ...formatting };
      const width = widths?.at(column);
      if (width === undefined) delete next.width;
      else next.width = width;
      return next;
    };
    const cells = [...row.cells];
    const original = { ...entry.cell, formatting: formattingForColumn(0) };
    const stamp = stamps.value[stampIndex++];
    if (stamp !== undefined)
      original.propertyChanges = [
        {
          type: "tableCellPropertyChange",
          info: stampInfo(stamp),
          ...(entry.cell.formatting === undefined
            ? {}
            : { previousFormatting: entry.cell.formatting }),
        },
      ];
    const mergeStamp = vertical ? stamps.value[stampIndex++] : undefined;
    const vMerge = entry.cell.formatting?.vMerge;
    if (mergeStamp !== undefined)
      original.structuralChange = {
        type: "tableCellMerge",
        info: stampInfo(mergeStamp),
        ...(vMerge === undefined
          ? {}
          : {
              verticalMergeOriginal:
                vMerge === CELL_VERTICAL_MERGES.RESTART ? "rest" : "continue",
            }),
      };
    const replacement: TableCell[] = [original];
    for (let column = 1; column < span; column += 1) {
      const cell = fresh.value[freshIndex++];
      if (cell !== undefined) {
        const next = { ...cell, formatting: formattingForColumn(column) };
        const cellStamp = stamps.value[stampIndex++];
        if (cellStamp !== undefined)
          next.structuralChange = {
            type: "tableCellInsertion",
            info: stampInfo(cellStamp),
          };
        replacement.push(next);
      }
    }
    cells.splice(entry.index, 1, ...replacement);
    return { ...row, cells };
  });
  return Result.ok({ ...table, rows });
};

/**
 * The grid a table edit works on: one the table's column widths name every
 * column of. A row that ends early omits the columns after its cells, as a
 * stated `w:gridAfter` does.
 */
const statedGrid = (
  table: Table,
  op: { type: DocumentOpType },
): Result<TableGrid, DocumentOpRefusal> =>
  tableGrid(table, op.type).andThen((grid) =>
    table.columnWidths === undefined || table.columnWidths.length === grid.width
      ? Result.ok(grid)
      : Result.err(
          new DocumentOpRefusal({
            opType: op.type,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
            message:
              "Column edits need every row to cover the table's column widths.",
          }),
        ),
  );

/** Count identities from the same topology/selection decisions the operation applies. */
export const tableEditParagraphDemand = (
  location: TableRowLocation,
  op: TableIntentOperation,
): Result<number, DocumentOpRefusal> => {
  if (isPropertyEdit(op)) return Result.ok(0);
  const table = location.table;
  const grid = statedGrid(table, op);
  if (grid.isErr()) return Result.err(grid.error);
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
      return Result.ok(
        grid.value.rows.filter((cells, row) => {
          const before = table.rows[row]?.formatting?.gridBefore ?? 0;
          const { end } = rowExtent(table, grid.value, row);
          return (
            op.column >= before &&
            op.column <= end &&
            !cells.some(
              (cell) => cell.start < op.column && op.column < cell.end,
            )
          );
        }).length,
      );
    case DOCUMENT_OP_TYPES.MERGE_CELLS: {
      const selection = mergeSelection({ table, grid: grid.value, op });
      if (selection.isErr()) return Result.err(selection.error);
      return Result.ok(
        selection.value.verticalOnly ? 0 : op.bottom - op.top - 1,
      );
    }
    case DOCUMENT_OP_TYPES.SPLIT_CELL: {
      const target = grid.value.rows[location.rowIndex]?.find(
        (entry) => entry.index === location.cellIndex,
      );
      if (!target) return mismatch(op, "The addressed cell is absent.");
      if (target.row !== target.ownerRow || target.index !== target.ownerIndex)
        return mismatch(
          op,
          "A vertical merge continuation must be addressed through its owner cell.",
        );
      const group = grid.value.rows.flatMap((row) =>
        row.filter(
          (cell) =>
            cell.ownerRow === target.ownerRow &&
            cell.ownerIndex === target.ownerIndex,
        ),
      );
      const span = target.end - target.start;
      if (span === 1 && group.length === 1)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE,
          message: "A one-row, one-column cell has nothing to split.",
        });
      const height = group.length;
      return Result.ok(height * (span - 1));
    }
    default:
      return Result.ok(0);
  }
};

type PropertyEditOptions = {
  op: PropertyEdit;
  table: Table;
  location: TableRowLocation;
};
const editProperties = ({
  op,
  table,
  location,
}: PropertyEditOptions): Result<Table, DocumentOpRefusal> => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.SET_TABLE_PROPS: {
      if (
        Object.keys(op.patch).some(
          (key) => !PROPERTY_FIELDS.table.some((field) => field === key),
        )
      )
        return mismatch(
          op,
          "Grid and source provenance require a grid operation.",
        );
      // One record per element: a further change keeps what the first recorded.
      if ((table.propertyChanges?.length ?? 0) > 1)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
          message: "Resolve the existing table property revision.",
        });
      const pending = table.propertyChanges?.at(0);
      const next = { ...table };
      const formatting = applyFormattingPatch(table.formatting, op.patch);
      if (!patchChanges(table.formatting, op.patch)) return Result.ok(table);
      if (formatting === undefined) delete next.formatting;
      else next.formatting = authored(formatting);
      if (op.revision !== undefined) {
        // A pending change already holds the formatting before any of them.
        const previousFormatting =
          pending === undefined ? table.formatting : pending.previousFormatting;
        next.propertyChanges = [
          {
            type: "tablePropertyChange",
            info: stampInfo(op.revision),
            ...(previousFormatting === undefined ? {} : { previousFormatting }),
          },
        ];
      }
      return Result.ok(next);
    }
    case DOCUMENT_OP_TYPES.SET_ROW_PROPS: {
      if (
        Object.keys(op.patch).some(
          (key) => !PROPERTY_FIELDS.row.some((field) => field === key),
        )
      )
        return mismatch(op, "Omitted grid slots require a topology operation.");
      const row = table.rows[location.rowIndex];
      if (!row) return mismatch(op, "The addressed row is absent.");
      // One record per element: a further change keeps what the first recorded.
      if ((row.propertyChanges?.length ?? 0) > 1)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
          message: "Resolve the existing row property revision.",
        });
      const pending = row.propertyChanges?.at(0);
      const next = { ...row };
      const formatting = applyFormattingPatch(row.formatting, op.patch);
      if (!patchChanges(row.formatting, op.patch)) return Result.ok(table);
      if (formatting === undefined) delete next.formatting;
      else next.formatting = authored(formatting);
      if (op.revision !== undefined) {
        // A pending change already holds the formatting before any of them.
        const previousFormatting =
          pending === undefined ? row.formatting : pending.previousFormatting;
        next.propertyChanges = [
          {
            type: "tableRowPropertyChange",
            info: stampInfo(op.revision),
            ...(previousFormatting === undefined ? {} : { previousFormatting }),
          },
        ];
      }
      const rows = [...table.rows];
      rows[location.rowIndex] = next;
      return Result.ok({ ...table, rows });
    }
    case DOCUMENT_OP_TYPES.SET_CELL_PROPS: {
      if (
        Object.keys(op.patch).some(
          (key) => !PROPERTY_FIELDS.cell.some((field) => field === key),
        )
      )
        return mismatch(op, "Cell spans require a merge or split operation.");
      const row = table.rows[location.rowIndex];
      const cell = row?.cells[location.cellIndex];
      if (!row || !cell) return mismatch(op, "The addressed cell is absent.");
      // One record per element: a further change keeps what the first recorded.
      if ((cell.propertyChanges?.length ?? 0) > 1)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
          message: "Resolve the existing cell property revision.",
        });
      const pending = cell.propertyChanges?.at(0);
      const next = { ...cell };
      const formatting = applyFormattingPatch(cell.formatting, op.patch);
      if (!patchChanges(cell.formatting, op.patch)) return Result.ok(table);
      if (formatting === undefined) delete next.formatting;
      else next.formatting = authored(formatting);
      if (op.revision !== undefined) {
        // A pending change already holds the formatting before any of them.
        const previousFormatting =
          pending === undefined ? cell.formatting : pending.previousFormatting;
        next.propertyChanges = [
          {
            type: "tableCellPropertyChange",
            info: stampInfo(op.revision),
            ...(previousFormatting === undefined ? {} : { previousFormatting }),
          },
        ];
      }
      const rows = [...table.rows];
      const cells = [...row.cells];
      cells[location.cellIndex] = next;
      rows[location.rowIndex] = { ...row, cells };
      return Result.ok({ ...table, rows });
    }
    default: {
      const exhaustive: never = op;
      return exhaustive;
    }
  }
};

/** Structural edits validate their grid; property edits retain every structural fact. */
export const applyTableEdit = (
  document: Document,
  op: Op,
): Result<DocumentEdit, DocumentOpRefusal> => {
  const located = locateTableRow(document, op);
  if (located.isErr()) return Result.err(located.error);
  const location = located.value;
  const table = location.table;
  if (op.type === DOCUMENT_OP_TYPES.SET_TABLE) {
    if (!equalForStaleness(table, op.expected))
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.STALE,
        message: "The table changed since this inverse was recorded.",
      });
    return commitTable({ document, op, location, table: op.table });
  }
  if (op.revision !== undefined) {
    const ambiguous = (record: object) =>
      ["formatting", "propertyChanges", "structuralChange"].some(
        (key) =>
          Object.hasOwn(record, key) && Reflect.get(record, key) === undefined,
      );
    if (
      ambiguous(table) ||
      table.rows.some((row) => ambiguous(row) || row.cells.some(ambiguous))
    ) {
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        message:
          "Explicit undefined review fields have no persisted OOXML property snapshot.",
      });
    }
  }
  if (isPropertyEdit(op)) {
    return editProperties({ op, table, location }).andThen((next) => {
      if (structurallyEqual(table, next))
        return Result.ok({
          document,
          inverse: [],
          touched: { modified: [], inserted: [], removed: [] },
        });
      return commitTable({ document, op, location, table: next });
    });
  }
  const grid = statedGrid(table, op);
  if (grid.isErr()) return Result.err(grid.error);
  // Removing the column that holds every cell the table has, as its only
  // column does, removes the table.
  if (
    op.type === DOCUMENT_OP_TYPES.DELETE_COLUMN &&
    op.column >= 0 &&
    op.column < grid.value.width &&
    grid.value.rows.every((entries) =>
      entries.every(
        ({ start, end }) => start === op.column && end === op.column + 1,
      ),
    )
  ) {
    const cells = table.rows.flatMap((row) => row.cells);
    if (cells.some(reachesPastRemovedCell) || strandsABookmark(cells))
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.DEPENDENT_RECORDS,
        message:
          "Resolve dependent anchors and review records before removing the table.",
      });
    return applyTableOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_TABLE,
      story: op.story,
      blockId: op.blockId,
      ...(op.revision === undefined ? {} : { revision: op.revision }),
      ...(op.newIds === undefined ? {} : { newIds: op.newIds }),
    }).mapError(
      (error) =>
        new DocumentOpRefusal({
          opType: op.type,
          reason: error.reason,
          message: error.message,
        }),
    );
  }
  let changed: Result<Table, DocumentOpRefusal>;
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
    case DOCUMENT_OP_TYPES.DELETE_COLUMN:
      changed = columnEdit({ document, op, table, grid: grid.value });
      break;
    case DOCUMENT_OP_TYPES.MERGE_CELLS:
      changed = merge({ document, op, table, grid: grid.value });
      break;
    case DOCUMENT_OP_TYPES.SPLIT_CELL:
      changed = split({ document, op, table, grid: grid.value, location });
      break;
    case DOCUMENT_OP_TYPES.SET_TABLE_GRID: {
      if (
        op.columnWidths.length !== grid.value.width ||
        op.columnWidths.some(
          (width) => !Number.isSafeInteger(width) || width <= 0,
        )
      )
        return mismatch(
          op,
          "Column widths must match the grid and be positive twip counts.",
        );
      if (table.formatting?.gridChange !== undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
          message: "Resolve the existing grid revision before resizing.",
        });
      if (op.revision !== undefined && table.columnWidths === undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          message: "A grid revision requires an explicit previous grid.",
        });
      if (table.propertyChanges !== undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
          message:
            "Resolve the table property history before resizing its grid.",
        });
      const stamps = stampsFor(op, 2);
      if (stamps.isErr()) return Result.err(stamps.error);
      const next = { ...table, columnWidths: [...op.columnWidths] };
      const gridStamp = stamps.value.at(0);
      const propertyStamp = stamps.value.at(1);
      if (gridStamp !== undefined && propertyStamp !== undefined) {
        next.formatting = {
          ...table.formatting,
          gridChange: {
            id: gridStamp.id,
            columnWidths: [...(table.columnWidths ?? [])],
          },
        };
        next.propertyChanges = [
          {
            type: "tablePropertyChange",
            info: stampInfo(propertyStamp),
            ...(table.formatting === undefined
              ? {}
              : { previousFormatting: table.formatting }),
          },
        ];
      }
      changed = Result.ok(next);
      break;
    }
    default: {
      const exhaustive: never = op;
      return exhaustive;
    }
  }
  if (changed.isErr()) return Result.err(changed.error);
  if (structurallyEqual(table, changed.value))
    return Result.ok({
      document,
      inverse: [],
      touched: { modified: [], inserted: [], removed: [] },
    });
  return commitTable({ document, op, location, table: changed.value });
};
