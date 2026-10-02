/** Semantic grid/cell/property edits; one exact, stale-checked table inverse. */
import { Result } from "better-result";
import {
  MAX_REVISION_ID,
  type Document,
  type Table,
  type TableCell,
  type TableMeasurement,
  type TableRow,
} from "../model/document";
import { storyBody, storyParagraphs, updateBlockList, withBodyContent } from "./blocks";
import { validateOpsDocument } from "./contract";
import type { DocumentEdit } from "./edits";
import { equalForStaleness, structurallyEqual } from "./equality";
import { idKey, isParaId, packageParagraphIds, paragraphIdsIn } from "./ids";
import { applyFormattingPatch } from "./patch";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { stampInfo } from "./review";
import { applyTableOp } from "./tables";
import { hasTableCellDependencies } from "./tableEditDependencies";
import { tableGrid, type TableGrid } from "./tableGrid";
import { locateTableRow, tableRowAnchor, type TableRowLocation } from "./tableLocation";
import {
  DOCUMENT_OP_TYPES,
  type TableEditOp,
  type TableIntentOperation,
  type SetTableOp,
  type RevisionStamp,
} from "./types";

const CELL_VERTICAL_MERGES = { RESTART: "restart", CONTINUE: "continue" } as const;

type Op = TableEditOp | SetTableOp;
type RefusalOptions = {
  op: Pick<Op, "type">;
  reason: DocumentOpRefusal["reason"];
  message: string;
};
const refuse = ({ op, reason, message }: RefusalOptions) =>
  Result.err(new DocumentOpRefusal({ opType: op.type, reason, message }));
const mismatch = (op: Pick<Op, "type">, message: string) =>
  refuse({ op, reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH, message });

const patchChanges = (base: object | undefined, patch: object): boolean =>
  Object.keys(patch).some((key) => {
    const value: unknown = Reflect.get(patch, key);
    if (value === undefined) return false;
    const previous: unknown = base === undefined ? undefined : Reflect.get(base, key);
    return value === null ? previous !== undefined : !structurallyEqual(previous, value);
  });

/** Source tokens describe the old property set and must not survive an authored edit. */
const authored = <T extends { sourceXml?: string }>(formatting: T): T => {
  const next = { ...formatting };
  delete next.sourceXml;
  return next;
};

type FreshCellsOptions = { document: Document; op: Op; ids: readonly string[]; count: number };
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

type CommitTableOptions = { document: Document; op: Op; location: TableRowLocation; table: Table };
const commitTable = ({
  document,
  op,
  location,
  table,
}: CommitTableOptions): Result<DocumentEdit, DocumentOpRefusal> => {
  const checked = tableGrid(table, op.type);
  if (checked.isErr()) return Result.err(checked.error);
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
    return mismatch(op, "Each cell must retain a paragraph outside its nested tables.");
  }
  const anchor = tableRowAnchor(table.rows);
  if (anchor === undefined) return mismatch(op, "The resulting table has no paragraph anchor.");
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
  const next = {
    ...document,
    package: { ...document.package, document: withBodyContent(body, content) },
  };
  const valid = validateOpsDocument(next);
  if (valid.isErr())
    return refuse({ op, reason: valid.error.reason, message: valid.error.message });
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
  if (!Number.isSafeInteger(first.value) || first.value < 0) return Result.err("invalid");
  let value = first.value;
  for (const cell of cells.slice(1)) {
    const width = cell.formatting?.width;
    // Unknown or incompatible widths do not add to a unit-bearing leading width.
    if (width === undefined || width.type !== first.type) return Result.ok(first);
    if (!Number.isSafeInteger(width.value) || width.value < 0) return Result.err("invalid");
    value += width.value;
    if (!Number.isSafeInteger(value)) return Result.err("invalid");
  }
  return Result.ok({ type: first.type, value });
};
const withOmission = (row: TableRow, side: "gridBefore" | "gridAfter", value: number): TableRow => {
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
const hasReview = (table: Table) =>
  table.formatting?.gridChange !== undefined ||
  table.propertyChanges !== undefined ||
  table.rows.some(
    (row) =>
      row.structuralChange !== undefined ||
      row.propertyChanges !== undefined ||
      row.cells.some(
        (cell) => cell.structuralChange !== undefined || cell.propertyChanges !== undefined,
      ),
  );

/** Multiple physical review records always get distinct, explicitly supplied ids. */
const stampsFor = (op: TableEditOp, count: number): Result<RevisionStamp[], DocumentOpRefusal> => {
  const revision = op.revision;
  if (revision === undefined) return Result.ok([]);
  const ids = [revision.id, ...(op.newIds?.revision ?? [])];
  if (new Set(ids).size !== ids.length)
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      message: "Revision ids must be distinct within the operation.",
    });
  if (ids.slice(0, count).some((id) => !Number.isInteger(id) || id < 0 || id > MAX_REVISION_ID))
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
  return Result.ok(ids.slice(0, count).map((id) => Object.assign({}, revision, { id })));
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
    return mismatch(op, "Deleting the final column requires deleting the table.");
  if (insert && (!Number.isSafeInteger(op.width) || op.width <= 0))
    return mismatch(op, "A column width must be a positive twip count.");
  if (hasStructuralMarkup(table))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "Indexed table markup must be relocated before changing columns.",
    });
  if (hasReview(table))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
      message: "Resolve table reviews before changing its columns.",
    });
  const boundaryRows = grid.rows.filter((cells, row) => {
    const before = table.rows[row]?.formatting?.gridBefore ?? 0;
    const after = grid.width - (table.rows[row]?.formatting?.gridAfter ?? 0);
    return (
      op.column >= before &&
      op.column <= after &&
      !cells.some((cell) => cell.start < op.column && op.column < cell.end)
    );
  }).length;
  const empty: TableCell[] = [];
  const fresh =
    op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN
      ? freshCells({ document, op, ids: op.newBlockIds, count: boundaryRows })
      : Result.ok(empty);
  if (fresh.isErr()) return Result.err(fresh.error);
  if (op.revision !== undefined && table.columnWidths === undefined) {
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "Tracked column edits require an explicit previous grid.",
    });
  }
  const stamps = stampsFor(op, table.rows.length + 2);
  if (stamps.isErr()) return Result.err(stamps.error);
  let freshIndex = 0;
  const rows: TableRow[] = [];
  for (const [rowIndex, row] of table.rows.entries()) {
    const entries = grid.rows[rowIndex];
    if (entries === undefined) return mismatch(op, "The grid row is missing.");
    const before = row.formatting?.gridBefore ?? 0;
    const after = grid.width - (row.formatting?.gridAfter ?? 0);
    if (op.column < before) {
      const next = withOmission(row, "gridBefore", before + (insert ? 1 : -1));
      const stamp = stamps.value[rowIndex];
      if (stamp !== undefined)
        next.propertyChanges = [
          {
            type: "tableRowPropertyChange",
            info: stampInfo(stamp),
            ...(row.formatting === undefined ? {} : { previousFormatting: row.formatting }),
          },
        ];
      rows.push(next);
      continue;
    }
    if (op.column >= after && (!insert || op.column > after)) {
      const next = withOmission(
        row,
        "gridAfter",
        (row.formatting?.gridAfter ?? 0) + (insert ? 1 : -1),
      );
      const stamp = stamps.value[rowIndex];
      if (stamp !== undefined)
        next.propertyChanges = [
          {
            type: "tableRowPropertyChange",
            info: stampInfo(stamp),
            ...(row.formatting === undefined ? {} : { previousFormatting: row.formatting }),
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
        const stamp = stamps.value[rowIndex];
        cells.push(
          stamp === undefined
            ? cell
            : { ...cell, structuralChange: { type: "tableCellInsertion", info: stampInfo(stamp) } },
        );
        inserted = true;
      }
      const hit = entry.start <= op.column && op.column < entry.end;
      if (insert && entry.start < op.column && hit) {
        const next = withSpan(entry.cell, entry.end - entry.start + 1);
        const stamp = stamps.value[rowIndex];
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
        if (entry.end - entry.start === 1 && hasTableCellDependencies(entry.cell))
          return refuse({
            op,
            reason: DOCUMENT_OP_REFUSAL_REASONS.DEPENDENT_RECORDS,
            message: "Resolve dependent anchors and review records before removing this cell.",
          });
        if (entry.end - entry.start > 1) {
          const next = withSpan(entry.cell, entry.end - entry.start - 1);
          const stamp = stamps.value[rowIndex];
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
        const stamp = stamps.value[rowIndex];
        if (stamp !== undefined)
          cells.push({
            ...entry.cell,
            structuralChange: { type: "tableCellDeletion", info: stampInfo(stamp) },
          });
        continue;
      }
      cells.push(entry.cell);
    }
    if (insert && !inserted) {
      const cell = fresh.value[freshIndex++];
      if (cell === undefined) return mismatch(op, "A fresh cell is missing.");
      const stamp = stamps.value[rowIndex];
      cells.push(
        stamp === undefined
          ? cell
          : { ...cell, structuralChange: { type: "tableCellInsertion", info: stampInfo(stamp) } },
      );
    }
    if (!cells.some((cell) => cell.structuralChange?.type !== "tableCellDeletion"))
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.TABLE_ROW_EMPTY,
        message: "Deleting this column would leave a row without an active cell.",
      });
    rows.push({ ...row, cells });
  }
  const columnWidths = table.columnWidths === undefined ? undefined : [...table.columnWidths];
  if (columnWidths !== undefined) {
    if (op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN) columnWidths.splice(op.column, 0, op.width);
    else columnWidths.splice(op.column, 1);
  }
  const next = { ...table, rows };
  if (columnWidths !== undefined) next.columnWidths = columnWidths;
  const gridStamp = stamps.value[table.rows.length];
  const propertyStamp = stamps.value[table.rows.length + 1];
  if (op.revision !== undefined && gridStamp !== undefined && propertyStamp !== undefined) {
    next.formatting = {
      ...table.formatting,
      gridChange: { id: gridStamp.id, columnWidths: [...(table.columnWidths ?? [])] },
    };
    next.propertyChanges = [
      {
        type: "tablePropertyChange",
        info: stampInfo(propertyStamp),
        ...(table.formatting === undefined ? {} : { previousFormatting: table.formatting }),
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
    .map((row) => row.filter((cell) => cell.start < op.right && cell.end > op.left));
  for (const row of selected) {
    if (
      row.at(0)?.start !== op.left ||
      row.at(-1)?.end !== op.right ||
      row.some((cell) => cell.start < op.left || cell.end > op.right || cell.ownerRow < op.top)
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
          cell.content.every((block) => block.type === "paragraph" && block.content.length === 0),
        ),
      );
  return Result.ok({ selected, verticalOnly });
};

type MergeOptions = {
  document: Document;
  op: Extract<TableEditOp, { type: "mergeCells" }>;
  table: Table;
  grid: TableGrid;
};
const merge = ({ document, op, table, grid }: MergeOptions): Result<Table, DocumentOpRefusal> => {
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
            rowIndex === op.top ? CELL_VERTICAL_MERGES.RESTART : CELL_VERTICAL_MERGES.CONTINUE,
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
                  entry.cell.formatting.vMerge === "restart" ? "rest" : "continue",
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
    return Result.ok({ ...table, rows });
  }
  if (op.revision !== undefined)
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message:
        "Cell merge markup cannot record horizontal content relocation; tracked moves are required.",
    });
  const fresh = freshCells({ document, op, ids: op.newBlockIds, count: op.bottom - op.top - 1 });
  if (fresh.isErr()) return Result.err(fresh.error);
  const first = selected.at(0)?.at(0);
  if (first === undefined) return mismatch(op, "No cells are selected.");
  const content = selected.flatMap((row) => row.flatMap(({ cell }) => cell.content));
  const mergedCellWidth = mergedWidth(selected.at(0)?.map(({ cell }) => cell) ?? []);
  if (mergedCellWidth.isErr())
    return mismatch(op, "Selected cell widths must be nonnegative safe integer units.");
  const width = mergedCellWidth.value;
  const rows = table.rows.map((row, index) => {
    const entries = selected[index - op.top];
    if (entries === undefined) return row;
    const leading = entries.at(0);
    if (leading === undefined) return row;
    const base = index === op.top ? { ...first.cell, content } : fresh.value[index - op.top - 1];
    if (base === undefined) return row;
    const cell = withSpan(base, op.right - op.left);
    const formatting = authored({ ...cell.formatting });
    if (index === op.top) {
      if (width === undefined) delete formatting.width;
      else formatting.width = width;
    }
    if (op.bottom - op.top > 1) formatting.vMerge = index === op.top ? "restart" : "continue";
    else delete formatting.vMerge;
    const cells = [...row.cells];
    cells.splice(leading.index, entries.length, { ...cell, formatting });
    return { ...row, cells };
  });
  return Result.ok({ ...table, rows });
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
  const target = grid.rows[location.rowIndex]?.find((entry) => entry.index === location.cellIndex);
  if (!target) return mismatch(op, "The addressed cell is absent.");
  if (target.row !== target.ownerRow || target.index !== target.ownerIndex)
    return mismatch(op, "A vertical merge continuation must be addressed through its owner cell.");
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
      (cell) => cell.ownerRow === target.ownerRow && cell.ownerIndex === target.ownerIndex,
    ),
  );
  if (group.length > 1)
    return mismatch(op, "Splitting a cell inside a vertical merge is unsupported.");
  const span = target.end - target.start;
  if (span === 1)
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE,
      message: "A one-cell, one-column split would not change the table.",
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
    return mismatch(op, "A split cell width must be a nonnegative safe integer unit.");
  const fresh = freshCells({ document, op, ids: op.newBlockIds, count: group.length * (span - 1) });
  if (fresh.isErr()) return Result.err(fresh.error);
  const stamps = stampsFor(op, group.length * span);
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
    const replacement: TableCell[] = [original];
    for (let column = 1; column < span; column += 1) {
      const cell = fresh.value[freshIndex++];
      if (cell !== undefined) {
        const next = { ...cell, formatting: formattingForColumn(column) };
        const cellStamp = stamps.value[stampIndex++];
        if (cellStamp !== undefined)
          next.structuralChange = { type: "tableCellInsertion", info: stampInfo(cellStamp) };
        replacement.push(next);
      }
    }
    cells.splice(entry.index, 1, ...replacement);
    return { ...row, cells };
  });
  return Result.ok({ ...table, rows });
};

/** Count identities from the same topology/selection decisions the operation applies. */
export const tableEditParagraphDemand = (
  location: TableRowLocation,
  op: TableIntentOperation,
): Result<number, DocumentOpRefusal> => {
  const table = location.table;
  const grid = tableGrid(table, op.type);
  if (grid.isErr()) return Result.err(grid.error);
  switch (op.type) {
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
      return Result.ok(
        grid.value.rows.filter((cells, row) => {
          const before = table.rows[row]?.formatting?.gridBefore ?? 0;
          const after = grid.value.width - (table.rows[row]?.formatting?.gridAfter ?? 0);
          return (
            op.column >= before &&
            op.column <= after &&
            !cells.some((cell) => cell.start < op.column && op.column < cell.end)
          );
        }).length,
      );
    case DOCUMENT_OP_TYPES.MERGE_CELLS: {
      const selection = mergeSelection({ table, grid: grid.value, op });
      if (selection.isErr()) return Result.err(selection.error);
      return Result.ok(selection.value.verticalOnly ? 0 : op.bottom - op.top - 1);
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
          (cell) => cell.ownerRow === target.ownerRow && cell.ownerIndex === target.ownerIndex,
        ),
      );
      if (group.length > 1)
        return mismatch(op, "Splitting a cell inside a vertical merge is unsupported.");
      const span = target.end - target.start;
      if (span === 1)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE,
          message: "A one-cell, one-column split would not change the table.",
        });
      const height = group.length;
      return Result.ok(height * (span - 1));
    }
    default:
      return Result.ok(0);
  }
};

/** Grid topology is validated before and after every semantic table edit. */
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
        (key) => Object.hasOwn(record, key) && Reflect.get(record, key) === undefined,
      );
    if (ambiguous(table) || table.rows.some((row) => ambiguous(row) || row.cells.some(ambiguous))) {
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        message: "Explicit undefined review fields have no persisted OOXML property snapshot.",
      });
    }
  }
  const grid = tableGrid(table, op.type);
  if (grid.isErr()) return Result.err(grid.error);
  if (op.type === DOCUMENT_OP_TYPES.DELETE_COLUMN && op.column === 0 && grid.value.width === 1) {
    if (table.rows.some((row) => row.cells.some(hasTableCellDependencies)))
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.DEPENDENT_RECORDS,
        message: "Resolve dependent anchors and review records before removing the table.",
      });
    return applyTableOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_TABLE,
      story: op.story,
      blockId: op.blockId,
      ...(op.revision === undefined ? {} : { revision: op.revision }),
      ...(op.newIds === undefined ? {} : { newIds: op.newIds }),
    }).mapError(
      (error) =>
        new DocumentOpRefusal({ opType: op.type, reason: error.reason, message: error.message }),
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
        op.columnWidths.some((width) => !Number.isSafeInteger(width) || width <= 0)
      )
        return mismatch(op, "Column widths must match the grid and be positive twip counts.");
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
          message: "Resolve the table property history before resizing its grid.",
        });
      const stamps = stampsFor(op, 2);
      if (stamps.isErr()) return Result.err(stamps.error);
      const next = { ...table, columnWidths: [...op.columnWidths] };
      const gridStamp = stamps.value.at(0);
      const propertyStamp = stamps.value.at(1);
      if (gridStamp !== undefined && propertyStamp !== undefined) {
        next.formatting = {
          ...table.formatting,
          gridChange: { id: gridStamp.id, columnWidths: [...(table.columnWidths ?? [])] },
        };
        next.propertyChanges = [
          {
            type: "tablePropertyChange",
            info: stampInfo(propertyStamp),
            ...(table.formatting === undefined ? {} : { previousFormatting: table.formatting }),
          },
        ];
      }
      changed = Result.ok(next);
      break;
    }
    case DOCUMENT_OP_TYPES.SET_TABLE_PROPS: {
      if (
        Object.keys(op.patch).some((key) =>
          ["sourceXml", "gridSourceXml", "gridChange", "preserved"].includes(key),
        )
      )
        return mismatch(op, "Grid and source provenance require a grid operation.");
      if (table.propertyChanges !== undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
          message: "Resolve the existing table property revision.",
        });
      const next = { ...table };
      const formatting = applyFormattingPatch(table.formatting, op.patch);
      if (!patchChanges(table.formatting, op.patch))
        return Result.ok({
          document,
          inverse: [],
          touched: { modified: [], inserted: [], removed: [] },
        });
      if (formatting === undefined) delete next.formatting;
      else next.formatting = authored(formatting);
      if (op.revision !== undefined)
        next.propertyChanges = [
          {
            type: "tablePropertyChange",
            info: stampInfo(op.revision),
            ...(table.formatting === undefined ? {} : { previousFormatting: table.formatting }),
          },
        ];
      changed = Result.ok(next);
      break;
    }
    case DOCUMENT_OP_TYPES.SET_ROW_PROPS: {
      if (
        Object.keys(op.patch).some((key) =>
          ["sourceXml", "gridBefore", "gridAfter", "preserved"].includes(key),
        )
      )
        return mismatch(op, "Omitted grid slots require a topology operation.");
      const row = table.rows[location.rowIndex];
      if (!row) return mismatch(op, "The addressed row is absent.");
      if (row.propertyChanges !== undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
          message: "Resolve the existing row property revision.",
        });
      const next = { ...row };
      const formatting = applyFormattingPatch(row.formatting, op.patch);
      if (!patchChanges(row.formatting, op.patch))
        return Result.ok({
          document,
          inverse: [],
          touched: { modified: [], inserted: [], removed: [] },
        });
      if (formatting === undefined) delete next.formatting;
      else next.formatting = authored(formatting);
      if (op.revision !== undefined)
        next.propertyChanges = [
          {
            type: "tableRowPropertyChange",
            info: stampInfo(op.revision),
            ...(row.formatting === undefined ? {} : { previousFormatting: row.formatting }),
          },
        ];
      const rows = [...table.rows];
      rows[location.rowIndex] = next;
      changed = Result.ok({ ...table, rows });
      break;
    }
    case DOCUMENT_OP_TYPES.SET_CELL_PROPS: {
      if (
        Object.keys(op.patch).some((key) =>
          ["sourceXml", "gridSpan", "vMerge", "preserved"].includes(key),
        )
      )
        return mismatch(op, "Cell spans require a merge or split operation.");
      const row = table.rows[location.rowIndex];
      const cell = row?.cells[location.cellIndex];
      if (!row || !cell) return mismatch(op, "The addressed cell is absent.");
      if (cell.propertyChanges !== undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
          message: "Resolve the existing cell property revision.",
        });
      const next = { ...cell };
      const formatting = applyFormattingPatch(cell.formatting, op.patch);
      if (!patchChanges(cell.formatting, op.patch))
        return Result.ok({
          document,
          inverse: [],
          touched: { modified: [], inserted: [], removed: [] },
        });
      if (formatting === undefined) delete next.formatting;
      else next.formatting = authored(formatting);
      if (op.revision !== undefined)
        next.propertyChanges = [
          {
            type: "tableCellPropertyChange",
            info: stampInfo(op.revision),
            ...(cell.formatting === undefined ? {} : { previousFormatting: cell.formatting }),
          },
        ];
      const rows = [...table.rows];
      const cells = [...row.cells];
      cells[location.cellIndex] = next;
      rows[location.rowIndex] = { ...row, cells };
      changed = Result.ok({ ...table, rows });
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
