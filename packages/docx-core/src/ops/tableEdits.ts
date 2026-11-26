/** Semantic grid/cell/property edits; one exact, stale-checked table inverse. */
import { panic, Result } from "better-result";

import {
  type BlockContent,
  type Document,
  finalizeParagraphFormatting,
  MAX_REVISION_ID,
  type Paragraph,
  type ParagraphContent,
  type Table,
  type TableCell,
  type TableCellFormatting,
  type TableFormatting,
  type TableMeasurement,
  type TableRow,
  type TableRowFormatting,
  type TableStructuralChangeInfo,
  type TrackedChangeInfo,
} from "../model/document";
import { maxOf } from "../model/extremes";
import {
  storyBody,
  storyParagraphs,
  updateBlockList,
  withBodyContent,
} from "./blocks";
import { validateOpsDocument } from "./contract";
import type { DocumentEdit } from "./edits";
import { equalForStaleness, structurallyEqual } from "./equality";
import { freshenIdentities, identityKeysOf } from "./identity";
import {
  IDENTITY_SPACES,
  idKey,
  isParaId,
  packageIdentityKeyIndex,
  packageParagraphIds,
  paragraphIdsIn,
  slotKey,
} from "./ids";
import { paragraphLength } from "./offsets";
import { applyFormattingPatch } from "./patch";
import { relocatePositionedMarkers } from "./positionedMarkers";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import { withTableRecordsResolved } from "./resolveTableRows";
import {
  foldedParagraphPropertyChange,
  reviewFieldsOf,
  stampInfo,
  withReviewFields,
  WRAP_KINDS,
  type WrapKind,
  wrapTracked,
} from "./review";
import { replaceStoryBody } from "./stories";
import {
  reachesPastRemovedCell,
  strandsABookmark,
} from "./tableEditDependencies";
import {
  editRemovesRow,
  type GridCell,
  type TableGrid,
  tableGrid,
} from "./tableGrid";
import {
  locateTableRow,
  tableRowAnchor,
  type TableRowLocation,
} from "./tableLocation";
import { rowsWithoutRow } from "./tableRowRemoval";
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
  // A cell narrowed to one column states no properties it did not have.
  if (Object.keys(formatting).length > 0) return { ...cell, formatting };
  const { formatting: _formatting, ...plain } = cell;
  return plain;
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

/** Captures and cells inside the same wrapper must move as one group. */
const hasWrappedCellMarkup = (table: Table): boolean => {
  const wrapped = ({
    contentControls,
    carrierStack,
  }: Pick<TableCell, "contentControls" | "carrierStack">) =>
    contentControls !== undefined || carrierStack !== undefined;
  return table.rows.some(
    (row) =>
      row.preserved?.children !== undefined &&
      (row.preserved.children.some(wrapped) || row.cells.some(wrapped)),
  );
};

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

/** The grid fields a cell merge or split rewrites, each with its unstated value. */
const CELL_GRID_FIELDS = {
  gridSpan: 1,
  vMerge: undefined,
  width: undefined,
} as const;
/** The cell field a column edit rewrites. */
const CELL_SPAN_FIELDS = { gridSpan: 1 } as const;
/** The row fields a column edit rewrites. */
const ROW_GRID_FIELDS = { gridBefore: 0, gridAfter: 0 } as const;
type GridFields =
  typeof CELL_GRID_FIELDS | typeof CELL_SPAN_FIELDS | typeof ROW_GRID_FIELDS;

const gridField = (
  formatting: object | undefined,
  field: string,
  fields: GridFields,
): unknown =>
  (formatting === undefined ? undefined : Reflect.get(formatting, field)) ??
  Reflect.get(fields, field);

/** Whether every pending property snapshot states the element's own grid fields. */
const statesOwnGrid = (
  changes: readonly { previousFormatting?: object }[] | undefined,
  formatting: object | undefined,
  fields: GridFields,
) =>
  (changes ?? []).every(({ previousFormatting }) =>
    Object.keys(fields).every((field) =>
      structurallyEqual(
        gridField(previousFormatting, field, fields),
        gridField(formatting, field, fields),
      ),
    ),
  );

type RewrittenFieldsOptions = {
  before: object | undefined;
  after: object | undefined;
  fields: GridFields;
};
const rewrittenFields = ({ before, after, fields }: RewrittenFieldsOptions) =>
  Object.keys(fields).filter(
    (field) =>
      !structurallyEqual(
        gridField(before, field, fields),
        gridField(after, field, fields),
      ),
  );

type RestateChangesOptions<C> = {
  changes: readonly C[] | undefined;
  before: object | undefined;
  after: object | undefined;
  fields: GridFields;
};
/**
 * Pending property changes over a direct edit's new grid: each snapshot takes
 * the span, merge, width or omission the edit gives the element, so resolving
 * it changes only the other properties. A snapshot stating another value of a
 * field the edit rewrites records an earlier grid: an error.
 */
const restateChanges = <
  C extends { previousFormatting?: object; currentFormatting?: object },
>({
  changes,
  before,
  after,
  fields,
}: RestateChangesOptions<C>): Result<C[] | undefined, "earlierGrid"> => {
  if (changes === undefined) return Result.ok(undefined);
  const rewritten = rewrittenFields({ before, after, fields });
  if (rewritten.length === 0) return Result.ok([...changes]);
  const restate = (
    snapshot: object | undefined,
  ): Result<Record<string, unknown> | undefined, "earlierGrid"> => {
    if (
      rewritten.some(
        (field) =>
          !structurallyEqual(
            gridField(snapshot, field, fields),
            gridField(before, field, fields),
          ),
      )
    )
      return Result.err("earlierGrid");
    const next: Record<string, unknown> = { ...snapshot };
    Reflect.deleteProperty(next, "sourceXml");
    for (const field of rewritten) {
      const value: unknown =
        after === undefined ? undefined : Reflect.get(after, field);
      if (value === undefined) Reflect.deleteProperty(next, field);
      else next[field] = value;
    }
    return Result.ok(
      snapshot === undefined && Object.keys(next).length === 0
        ? undefined
        : next,
    );
  };
  const restated: C[] = [];
  for (const change of changes) {
    const previous = restate(change.previousFormatting);
    if (previous.isErr()) return Result.err(previous.error);
    const next = { ...change };
    delete next.previousFormatting;
    if (previous.value !== undefined)
      Object.assign(next, { previousFormatting: previous.value });
    if (change.currentFormatting !== undefined) {
      const current = restate(change.currentFormatting);
      if (current.isErr()) return Result.err(current.error);
      delete next.currentFormatting;
      if (current.value !== undefined)
        Object.assign(next, { currentFormatting: current.value });
    }
    restated.push(next);
  }
  return Result.ok(restated);
};

/** `after`, the direct edit of `before`, with `before`'s pending property changes restated. */
const restatedCell = (
  before: TableCell,
  after: TableCell,
  fields: GridFields,
): Result<TableCell, "earlierGrid"> => {
  const changes = before.propertyChanges;
  if (changes === undefined) return Result.ok(after);
  // A snapshot of the cell's own earlier insertion, deletion or merge would
  // come back on the edited cell when rejected.
  if (
    changes.some(
      ({ previousStructuralChange }) => previousStructuralChange !== undefined,
    ) &&
    rewrittenFields({
      before: before.formatting,
      after: after.formatting,
      fields,
    }).length > 0
  )
    return Result.err("earlierGrid");
  return restateChanges({
    changes,
    before: before.formatting,
    after: after.formatting,
    fields,
  }).map((propertyChanges) =>
    propertyChanges === undefined
      ? after
      : Object.assign({}, after, { propertyChanges }),
  );
};

type CopiedPropertyChangesOptions = {
  /** The cell whose pending property changes a new cell takes. */
  source: TableCell;
  /** The new cell, with the properties it takes from `source`. */
  cell: TableCell;
  /** The ids of the new records, drawn one for each change. */
  ids: Iterator<number>;
};
/**
 * A new cell a direct edit makes from `source`, taking its properties, with
 * `source`'s pending property changes restated over them as records of its
 * own: rejecting them takes the other properties off the new cell as it
 * does off `source`.
 */
const copiedPropertyChanges = ({
  source,
  cell,
  ids,
}: CopiedPropertyChangesOptions): Result<TableCell, "earlierGrid"> => {
  const restated = restatedCell(source, cell, CELL_GRID_FIELDS);
  if (restated.isErr() || restated.value.propertyChanges === undefined)
    return restated;
  const propertyChanges: NonNullable<TableCell["propertyChanges"]> = [];
  for (const change of restated.value.propertyChanges) {
    const next = ids.next();
    if (next.done === true)
      panic("A copied property change has no revision id.");
    propertyChanges.push({
      ...change,
      info: { ...change.info, id: next.value },
    });
  }
  return Result.ok(Object.assign({}, restated.value, { propertyChanges }));
};

type DirectRevisionIdsOptions = {
  document: Document;
  op: TableEditOp;
  count: number;
};
/**
 * The revision ids a direct edit gives the records it adds beside pending
 * ones, such as a new cell's copy of its source's property change: the first
 * `count` ids of `newIds.revision` no record in the package holds.
 */
const directRevisionIds = ({
  document,
  op,
  count,
}: DirectRevisionIdsOptions): Result<number[], DocumentOpRefusal> => {
  if (count === 0) return Result.ok([]);
  const held = packageIdentityKeyIndex(document.package);
  const ids: number[] = [];
  for (const id of op.newIds?.revision ?? []) {
    if (ids.length === count) break;
    if (!Number.isInteger(id) || id < 0 || id > MAX_REVISION_ID)
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        message: "Revision ids must be nonnegative 31-bit integers.",
      });
    if (
      ids.includes(id) ||
      held.has(slotKey({ space: IDENTITY_SPACES.REVISION, id }))
    )
      continue;
    ids.push(id);
  }
  if (ids.length < count)
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
      message: `This table change keeps pending records under ${count} new revision ids.`,
    });
  return Result.ok(ids);
};

/** `after`, the direct edit of `before`, with `before`'s pending property changes restated. */
const restatedRow = (
  before: TableRow,
  after: TableRow,
): Result<TableRow, "earlierGrid"> =>
  restateChanges({
    changes: before.propertyChanges,
    before: before.formatting,
    after: after.formatting,
    fields: ROW_GRID_FIELDS,
  }).map((propertyChanges) =>
    propertyChanges === undefined
      ? after
      : Object.assign({}, after, { propertyChanges }),
  );

const isCellInsertionOrDeletion = (cell: TableCell) =>
  cell.structuralChange?.type === "tableCellInsertion" ||
  cell.structuralChange?.type === "tableCellDeletion";

/**
 * The grid columns a pending insertion or deletion takes whole: every cell
 * standing in the column is a one-column cell under one kind of pending
 * insertion or deletion. Undefined when such a change takes only part of a
 * column, which leaves no column a direct edit can place itself against.
 */
const reviewColumns = (grid: TableGrid) => {
  const inserted = new Set<number>();
  for (let column = 0; column < grid.width; column += 1) {
    const standing = grid.rows.flatMap((entries) =>
      entries.filter(({ start, end }) => start <= column && column < end),
    );
    const kinds = new Set(
      standing.map(({ cell, start, end }) =>
        end - start === 1 && isCellInsertionOrDeletion(cell)
          ? cell.structuralChange?.type
          : undefined,
      ),
    );
    if (kinds.size > 1) return undefined;
    if (kinds.has("tableCellInsertion")) inserted.add(column);
  }
  const whole = grid.rows.every((entries) =>
    entries.every(
      ({ cell, start, end }) =>
        !isCellInsertionOrDeletion(cell) || end - start === 1,
    ),
  );
  return whole ? { inserted } : undefined;
};

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

type ColumnReviewsOptions = {
  op: Extract<TableEditOp, { type: "insertColumn" | "deleteColumn" }>;
  table: Table;
  grid: TableGrid;
};
/**
 * Whether a direct column edit keeps the table's pending records, and the
 * earlier grid a pending grid change holds with the edit applied to it, so
 * rejecting the change restores the earlier widths around the edited column.
 * A grid column stands in the earlier grid unless a pending insertion takes
 * it whole.
 */
const columnReviews = ({
  op,
  table,
  grid,
}: ColumnReviewsOptions): Result<
  TableFormatting["gridChange"],
  DocumentOpRefusal
> => {
  const conflict = (message: string) =>
    refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
      message,
    });
  // Records taking part of a column, or recording other spans or
  // omissions, take `columnEditAcrossRecords`.
  const columns =
    reviewColumns(grid) ??
    panic("A column edit across partial records reads both decisions.");
  const gridChange = table.formatting?.gridChange;
  if (gridChange === undefined) return Result.ok(undefined);
  if (gridChange.columnWidths.length !== grid.width - columns.inserted.size)
    return conflict(
      "Resolve the pending grid change before changing the table's columns.",
    );
  const earlier =
    op.column -
    [...columns.inserted].filter((column) => column < op.column).length;
  const columnWidths = [...gridChange.columnWidths];
  if (op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN) {
    // A grid stating no widths at all stays one.
    const width = columnWidths.every((stated) => stated === undefined)
      ? undefined
      : op.width;
    columnWidths.splice(
      earlier,
      0,
      ...Array.from({ length: op.count ?? 1 }, () => width),
    );
  } else if (!columns.inserted.has(op.column)) columnWidths.splice(earlier, 1);
  return Result.ok({ ...gridChange, columnWidths });
};

type ColumnOp = Extract<TableEditOp, { type: "insertColumn" | "deleteColumn" }>;

/**
 * Whether a pending record places a column differently once it is accepted
 * than once it is rejected: a cell insertion or deletion that takes only part
 * of a column, or a row or cell property change recording other spans or
 * omissions.
 */
const recordsPlaceColumnsApart = (table: Table, grid: TableGrid) =>
  reviewColumns(grid) === undefined ||
  table.rows.some(
    (row) =>
      !statesOwnGrid(row.propertyChanges, row.formatting, ROW_GRID_FIELDS) ||
      row.cells.some(
        (cell) =>
          !statesOwnGrid(
            cell.propertyChanges,
            cell.formatting,
            CELL_SPAN_FIELDS,
          ),
      ),
  );

/** A cell's first paragraph id, which names it in every decision's table. */
const cellKey = (cell: TableCell): string | undefined =>
  storyParagraphs({ content: cell.content }).at(0)?.paragraph.paraId;

/** A row's cells in one decision's table, by key, and the columns it omits. */
type DecidedRow = {
  before: number;
  end: number;
  after: number;
  cells: readonly { key: string; start: number; end: number }[];
};

/** The rows of the table every cell and property record of which is decided. */
const decidedRows = (
  table: Table,
  reject: boolean,
  op: ColumnOp,
): DecidedRow[] | undefined => {
  const decided = withTableRecordsResolved(table, reject);
  if (decided.isErr() || decided.value.rows.length !== table.rows.length)
    return undefined;
  const grid = tableGrid(decided.value, op.type);
  if (grid.isErr()) return undefined;
  const rows: DecidedRow[] = [];
  for (const [index, row] of decided.value.rows.entries()) {
    const { before, end } = rowExtent(decided.value, grid.value, index);
    const cells = [];
    for (const { cell, start, end: cellEnd } of grid.value.rows[index] ?? []) {
      const key = cellKey(cell);
      if (key === undefined) return undefined;
      cells.push({ key, start, end: cellEnd });
    }
    rows.push({ before, end, after: row.formatting?.gridAfter ?? 0, cells });
  }
  return rows;
};

/** Where a column edit lands in one row of one decision's table. */
type ColumnPlace =
  | { type: "leaves" }
  | { type: "before" }
  | { type: "after" }
  /** An insertion between two cells, after the keyed one, or first. */
  | { type: "between"; after: string | undefined }
  /** The keyed cell spans the column: an insertion widens it, a deletion narrows or removes it. */
  | { type: "within"; cell: string; span: number };

const columnPlace = (row: DecidedRow, op: ColumnOp): ColumnPlace => {
  const insert = op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN;
  const column = op.column;
  if (column < row.before) return { type: "before" };
  if (insert ? column > row.end + row.after : column >= row.end + row.after)
    return { type: "leaves" };
  if (insert ? column > row.end : column >= row.end) return { type: "after" };
  const spanning = row.cells.find(({ start, end }) =>
    insert ? start < column && column < end : start <= column && column < end,
  );
  if (spanning !== undefined)
    return {
      type: "within",
      cell: spanning.key,
      span: spanning.end - spanning.start,
    };
  return {
    type: "between",
    after: row.cells.find(({ end }) => end === column)?.key,
  };
};

/** How one row's cells change, read against both decisions. */
type RowColumnPlan =
  | { type: "unchanged" }
  | { type: "omission"; side: "gridBefore" | "gridAfter" }
  | {
      type: "cells";
      /** Where new cells go, and which decision keeps them. */
      added: { at: number; kept: "both" | "accept" | "reject" } | undefined;
      /** The span each cell gains or loses, in the current and the earlier properties. */
      spans: ReadonlyMap<number, { current: number; earlier: number }>;
      /** The cells the edit removes. */
      removed: ReadonlySet<number>;
    };

type RowColumnPlanOptions = {
  entries: readonly GridCell[];
  accepted: DecidedRow;
  rejected: DecidedRow;
  op: ColumnOp;
};
/**
 * One row's part of a direct column edit, so that accepting every record
 * gives the edit of the accepted row, and rejecting them that of the
 * rejected row. `undefined` where no one change of the row does both.
 */
const rowColumnPlan = ({
  entries,
  accepted,
  rejected,
  op,
}: RowColumnPlanOptions): RowColumnPlan | undefined => {
  const inAccepted = accepted.cells.map(({ key }) => key);
  const inRejected = rejected.cells.map(({ key }) => key);
  const placeA = columnPlace(accepted, op);
  const placeR = columnPlace(rejected, op);
  if (placeA.type === "leaves" && placeR.type === "leaves")
    return { type: "unchanged" };
  if (placeA.type === "before" || placeA.type === "after") {
    if (placeR.type !== placeA.type) return undefined;
    return {
      type: "omission",
      side: placeA.type === "before" ? "gridBefore" : "gridAfter",
    };
  }
  // A row one decision leaves as it is takes the edit in the other only.
  if (placeR.type === "before" || placeR.type === "after") return undefined;
  const keys = entries.map(({ cell }) => cellKey(cell));
  const insert = op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN;
  const count =
    op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN ? (op.count ?? 1) : 1;
  const spans = new Map<number, { current: number; earlier: number }>();
  const removed = new Set<number>();
  // The cell standing in the deleted column as the row stands now.
  const occupant = insert
    ? undefined
    : entries.findIndex(
        ({ start, end }) => start <= op.column && op.column < end,
      );
  for (const [index, key] of keys.entries()) {
    if (key === undefined) return undefined;
    const presentA = inAccepted.includes(key);
    const presentR = inRejected.includes(key);
    const hitA = placeA.type === "within" && placeA.cell === key;
    const hitR = placeR.type === "within" && placeR.cell === key;
    if (!insert) {
      // A cell goes where every decision keeping it removes it.
      const goesA = hitA && placeA.span === 1;
      const goesR = hitR && placeR.span === 1;
      if (presentA && presentR && goesA !== goesR) return undefined;
      if (goesA || goesR) {
        removed.add(index);
        continue;
      }
    }
    const delta = insert ? count : -1;
    // A new cell only rejecting keeps is a pending deletion beside the cell
    // accepting widens: accepted, it hands that cell its columns. A column
    // accepting narrows a cell over that is a deleted cell's, as accepted it
    // was handed, goes with that cell.
    const current =
      hitA &&
      (insert
        ? placeR.type !== "between"
        : occupant === undefined || occupant === index)
        ? delta
        : 0;
    const earlier = hitR ? delta : 0;
    if (current === 0 && earlier === 0) continue;
    const cell = entries[index]?.cell;
    const recorded = (cell?.propertyChanges?.length ?? 0) > 0;
    // Without a record, both decisions read the cell's one span.
    if (!recorded && presentA && presentR && current !== earlier)
      return undefined;
    spans.set(
      index,
      recorded
        ? { current, earlier }
        : { current: presentA ? current : earlier, earlier: 0 },
    );
  }
  if (!insert) return { type: "cells", added: undefined, spans, removed };
  const range = (place: ColumnPlace, present: readonly string[]) => {
    if (place.type !== "between") return undefined;
    const lower = place.after === undefined ? 0 : keys.indexOf(place.after) + 1;
    const next = keys.findIndex(
      (key, index) =>
        index >= lower && key !== undefined && present.includes(key),
    );
    return { lower, upper: next === -1 ? keys.length : next };
  };
  const rangeA = range(placeA, inAccepted);
  const rangeR = range(placeR, inRejected);
  const ranges = [rangeA, rangeR].filter((each) => each !== undefined);
  if (ranges.length === 0)
    return { type: "cells", added: undefined, spans, removed };
  const at = maxOf(
    ranges.map(({ lower }) => lower),
    0,
  );
  if (ranges.some(({ lower, upper }) => lower < 0 || at > upper))
    return undefined;
  let kept: "both" | "accept" | "reject" = "both";
  if (rangeA === undefined) kept = "reject";
  else if (rangeR === undefined) kept = "accept";
  return { type: "cells", added: { at, kept }, spans, removed };
};

type ColumnPlanOptions = { table: Table; grid: TableGrid; op: ColumnOp };
/** Every row's part of a direct column edit across records that place it apart. */
const columnPlanAcrossRecords = ({
  table,
  grid,
  op,
}: ColumnPlanOptions): RowColumnPlan[] | undefined => {
  if (table.formatting?.gridChange !== undefined) return undefined;
  const accepted = decidedRows(table, false, op);
  const rejected = decidedRows(table, true, op);
  if (accepted === undefined || rejected === undefined) return undefined;
  const plans: RowColumnPlan[] = [];
  for (const index of table.rows.keys()) {
    const acceptedRow = accepted[index];
    const rejectedRow = rejected[index];
    const entries = grid.rows[index];
    if (
      acceptedRow === undefined ||
      rejectedRow === undefined ||
      entries === undefined
    )
      return undefined;
    const plan = rowColumnPlan({
      entries,
      accepted: acceptedRow,
      rejected: rejectedRow,
      op,
    });
    if (plan === undefined) return undefined;
    plans.push(plan);
  }
  return plans;
};

/** The span field of cell formatting, moved by `delta` columns. */
const respanned = (
  formatting: TableCellFormatting | undefined,
  delta: number,
): TableCellFormatting | undefined => {
  if (delta === 0) return formatting;
  const next = authored({ ...formatting });
  const span = (formatting?.gridSpan ?? 1) + delta;
  if (span === 1) delete next.gridSpan;
  else next.gridSpan = span;
  return next;
};

/** The omission field of row formatting, moved by `delta` columns. */
const reomitted = (
  formatting: TableRowFormatting | undefined,
  side: "gridBefore" | "gridAfter",
  delta: number,
): TableRowFormatting => {
  const next = authored({ ...formatting });
  const value = (formatting?.[side] ?? 0) + delta;
  if (side === "gridBefore") {
    if (value === 0) delete next.gridBefore;
    else next.gridBefore = value;
  } else if (value === 0) delete next.gridAfter;
  else next.gridAfter = value;
  return next;
};

type ColumnEditAcrossRecordsOptions = {
  document: Document;
  op: ColumnOp;
  table: Table;
  grid: TableGrid;
  conflict: (message: string) => Result<never, DocumentOpRefusal>;
};
/**
 * A direct column edit beside records that place the column apart once
 * accepted and once rejected. Each row takes the edit as both decisions read
 * it: a new cell only one of them keeps is recorded as that decision's cell
 * insertion or deletion, under the author of the row's record, and a span
 * only one of them changes is changed in the properties that decision gives
 * the cell. Accepting every record gives the edit of the accepted table, and
 * rejecting them that of the rejected one; where no table does both, the
 * edit refuses.
 */
const columnEditAcrossRecords = ({
  document,
  op,
  table,
  grid,
  conflict,
}: ColumnEditAcrossRecordsOptions): Result<Table, DocumentOpRefusal> => {
  const refused = () =>
    conflict(
      "Resolve the pending cell insertions, deletions and span changes that place the column apart before changing its columns.",
    );
  const plans = columnPlanAcrossRecords({ table, grid, op });
  if (plans === undefined) return refused();
  const count =
    op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN ? (op.count ?? 1) : 1;
  const addedRows = plans.filter(
    (plan) => plan.type === "cells" && plan.added !== undefined,
  );
  const fresh =
    op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN
      ? freshCells({
          document,
          op,
          ids: op.newBlockIds,
          count: addedRows.length * count,
        })
      : Result.ok([]);
  if (fresh.isErr()) return Result.err(fresh.error);
  const recorded = addedRows.filter(
    (plan) => plan.type === "cells" && plan.added?.kept !== "both",
  ).length;
  const ids = directRevisionIds({ document, op, count: recorded * count });
  if (ids.isErr()) return Result.err(ids.error);
  const freshIds = ids.value.values();
  const freshCellsLeft = fresh.value.values();
  const removedCells: TableCell[] = [];
  const rows: TableRow[] = [];
  for (const [rowIndex, row] of table.rows.entries()) {
    const plan = plans[rowIndex] ?? panic("A row has no plan.");
    const entries = grid.rows[rowIndex] ?? panic("A row has no grid.");
    const delta = op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN ? count : -1;
    switch (plan.type) {
      case "unchanged":
        rows.push(row);
        continue;
      case "omission": {
        const next = {
          ...row,
          formatting: reomitted(row.formatting, plan.side, delta),
        };
        if (row.propertyChanges !== undefined)
          next.propertyChanges = row.propertyChanges.map((change) => ({
            ...change,
            previousFormatting: reomitted(
              change.previousFormatting,
              plan.side,
              delta,
            ),
          }));
        rows.push(next);
        continue;
      }
      case "cells":
        break;
      default: {
        const unreachable: never = plan;
        return unreachable;
      }
    }
    // The record a new cell only one decision keeps is made under: the
    // row's first record, which places the column apart.
    const source =
      row.cells.find((cell) => cell.structuralChange !== undefined)
        ?.structuralChange?.info ??
      row.cells
        .find((cell) => cell.propertyChanges !== undefined)
        ?.propertyChanges?.at(0)?.info ??
      row.propertyChanges?.at(0)?.info;
    const recordInfo = (id: number): TrackedChangeInfo => {
      if (source === undefined) return panic("A new cell has no record.");
      const { author, date, initials, suggestionId } = source;
      return {
        id,
        author,
        ...(date === undefined ? {} : { date }),
        ...(initials === undefined ? {} : { initials }),
        ...(suggestionId === undefined ? {} : { suggestionId }),
      };
    };
    const cells: TableCell[] = [];
    const added = (): TableCell[] => {
      const out: TableCell[] = [];
      for (let column = 0; column < count; column += 1) {
        const next = freshCellsLeft.next();
        if (next.done === true) panic("A new cell has no paragraph id.");
        const kept = plan.added?.kept;
        if (kept === undefined || kept === "both") {
          out.push(next.value);
          continue;
        }
        const id = freshIds.next();
        if (id.done === true) panic("A new cell has no revision id.");
        out.push({
          ...next.value,
          structuralChange: {
            type:
              kept === "accept" ? "tableCellInsertion" : "tableCellDeletion",
            info: recordInfo(id.value),
          },
        });
      }
      return out;
    };
    for (const [index, { cell }] of entries.entries()) {
      if (plan.added?.at === index) cells.push(...added());
      if (plan.removed.has(index)) {
        removedCells.push(cell);
        continue;
      }
      const span = plan.spans.get(index);
      if (span === undefined) {
        cells.push(cell);
        continue;
      }
      const next: TableCell = { ...cell };
      const formatting = respanned(cell.formatting, span.current);
      if (formatting === undefined) delete next.formatting;
      else next.formatting = formatting;
      const [first, ...rest] = cell.propertyChanges ?? [];
      const earlier = respanned(first?.previousFormatting ?? {}, span.earlier);
      if (first !== undefined && span.earlier !== 0)
        next.propertyChanges = [
          {
            ...first,
            ...(earlier === undefined ? {} : { previousFormatting: earlier }),
          },
          ...rest,
        ];
      cells.push(next);
    }
    if (plan.added !== undefined && plan.added.at >= entries.length)
      cells.push(...added());
    if (editRemovesRow({ before: row.cells, after: cells })) continue;
    rows.push({ ...row, cells });
  }
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
  const columnWidths =
    table.columnWidths === undefined ? undefined : [...table.columnWidths];
  if (columnWidths !== undefined) {
    if (op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN)
      columnWidths.splice(
        op.column,
        0,
        ...Array.from({ length: count }, () => op.width),
      );
    else columnWidths.splice(op.column, 1);
  }
  const edited: Table = { ...table, rows };
  if (columnWidths !== undefined) edited.columnWidths = columnWidths;
  // Each decision of the edited table is the edit of that decision's table.
  const newIds =
    op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN ? op.newBlockIds : [];
  const unnamed = (value: Table): unknown => {
    let text = JSON.stringify(value);
    for (const id of newIds) text = text.replaceAll(`"${id}"`, `"new"`);
    const parsed: unknown = JSON.parse(text);
    return parsed;
  };
  for (const reject of [false, true]) {
    const decided = withTableRecordsResolved(table, reject);
    const editedDecided = withTableRecordsResolved(edited, reject);
    if (decided.isErr() || editedDecided.isErr()) return refused();
    const decidedGrid = statedGrid(decided.value, op);
    if (decidedGrid.isErr()) return refused();
    const want = columnEdit({
      document,
      op:
        op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN
          ? {
              ...op,
              newBlockIds: newIds.slice(
                0,
                plans.filter(
                  (plan) =>
                    plan.type === "cells" &&
                    plan.added !== undefined &&
                    plan.added.kept !== (reject ? "accept" : "reject"),
                ).length * count,
              ),
            }
          : op,
      table: decided.value,
      grid: decidedGrid.value,
    });
    if (
      want.isErr() ||
      !structurallyEqual(unnamed(editedDecided.value), unnamed(want.value))
    )
      return refused();
  }
  return Result.ok(edited);
};

type ColumnOptions = {
  document: Document;
  op: ColumnOp;
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
  const count =
    op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN ? (op.count ?? 1) : 1;
  if (!Number.isSafeInteger(count) || count < 1)
    return mismatch(op, "A column insertion inserts at least one column.");
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
  // A table records one previous grid (`w:tblGridChange`). A second tracked
  // insertion would have to share the pending one's record, and rejecting
  // either alone could not give back the grid without its own columns: the
  // columns of one change are inserted by one operation (`count`).
  if (
    insert &&
    op.revision !== undefined &&
    table.formatting?.gridChange !== undefined
  )
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_GRID_CHANGE,
      message:
        "The table's columns already have a pending change; insert every column of one change in one operation.",
    });
  // Another suggested column deletion may join pending ones: each keeps its
  // own cells and grid entry until it is resolved.
  const suggestedDeletion = !insert && op.revision !== undefined;
  if (
    op.revision !== undefined &&
    hasReviewBlockingColumns(table, suggestedDeletion)
  )
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
      message: "Resolve table reviews before changing its columns.",
    });
  // Records placing the column apart once accepted and once rejected: the
  // edit is read against both decisions.
  if (op.revision === undefined && recordsPlaceColumnsApart(table, grid))
    return columnEditAcrossRecords({
      document,
      op,
      table,
      grid,
      conflict: (message) =>
        refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
          message,
        }),
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
  // The rows an insertion gives new cells: those with a cell boundary at the
  // column (a cell spanning it widens instead).
  const takesNewCells = (row: number) => {
    const { before, end } = rowExtent(table, grid, row);
    return (
      insert &&
      op.column >= before &&
      op.column <= end &&
      !(grid.rows[row] ?? []).some(
        (cell) => cell.start < op.column && op.column < cell.end,
      )
    );
  };
  const boundaryRows = grid.rows.filter((_cells, row) =>
    takesNewCells(row),
  ).length;
  const empty: TableCell[] = [];
  const fresh =
    op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN
      ? freshCells({
          document,
          op,
          ids: op.newBlockIds,
          count: boundaryRows * count,
        })
      : Result.ok(empty);
  if (fresh.isErr()) return Result.err(fresh.error);
  // A suggested deletion keeps the grid while it is pending: the deleted
  // cells stay in their columns, and accepting removes them with their grid
  // entries.
  const trackedDeletion = suggestedDeletion && removedCells.length > 0;
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
  const earlierGrid =
    op.revision === undefined
      ? columnReviews({ op, table, grid })
      : Result.ok(undefined);
  if (earlierGrid.isErr()) return Result.err(earlierGrid.error);
  // Each row the edit changes takes a stamp, one per new cell where it gets
  // new cells; an edit besides a suggested deletion stamps the grid and
  // table properties after them.
  const stampedRows = new Map<number, number>();
  let rowStamps = 0;
  for (const rowIndex of table.rows.keys()) {
    if (leavesRow(rowIndex)) continue;
    stampedRows.set(rowIndex, rowStamps);
    rowStamps += takesNewCells(rowIndex) ? count : 1;
  }
  const stamps = stampsFor(op, trackedDeletion ? rowStamps : rowStamps + 2);
  if (stamps.isErr()) return Result.err(stamps.error);
  const rowStamp = (rowIndex: number, cell = 0) => {
    const first = stampedRows.get(rowIndex);
    return first === undefined ? undefined : stamps.value[first + cell];
  };
  // A direct edit keeps the pending property changes of what it re-spans,
  // over the spans and omissions it gives.
  const earlierSpans = () =>
    new DocumentOpRefusal({
      opType: op.type,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
      message:
        "A pending change records another span than the column edit changes.",
    });
  const keptRow = (before: TableRow, after: TableRow) =>
    op.revision === undefined
      ? restatedRow(before, after).mapError(earlierSpans)
      : Result.ok(after);
  const keptCell = (before: TableCell, after: TableCell) =>
    op.revision === undefined
      ? restatedCell(before, after, CELL_SPAN_FIELDS).mapError(earlierSpans)
      : Result.ok(after);
  let freshIndex = 0;
  /** The row's new cells, each recorded under a stamp of its own. */
  const newCells = (
    rowIndex: number,
  ): Result<TableCell[], DocumentOpRefusal> => {
    const cells: TableCell[] = [];
    for (let column = 0; column < count; column++) {
      const cell = fresh.value[freshIndex++];
      if (cell === undefined) return mismatch(op, "A fresh cell is missing.");
      const stamp = rowStamp(rowIndex, column);
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
    return Result.ok(cells);
  };
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
      const next = withOmission(
        row,
        "gridBefore",
        before + (insert ? count : -1),
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
      const kept = keptRow(row, next);
      if (kept.isErr()) return Result.err(kept.error);
      rows.push(kept.value);
      continue;
    }
    if (op.column >= end && (!insert || op.column > end)) {
      const next = withOmission(
        row,
        "gridAfter",
        (row.formatting?.gridAfter ?? 0) + (insert ? count : -1),
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
      const kept = keptRow(row, next);
      if (kept.isErr()) return Result.err(kept.error);
      rows.push(kept.value);
      continue;
    }
    const cells: TableCell[] = [];
    let inserted = false;
    for (const entry of entries) {
      if (insert && entry.start === op.column) {
        const added = newCells(rowIndex);
        if (added.isErr()) return Result.err(added.error);
        cells.push(...added.value);
        inserted = true;
      }
      const hit = entry.start <= op.column && op.column < entry.end;
      if (insert && entry.start < op.column && hit) {
        const next = withSpan(entry.cell, entry.end - entry.start + count);
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
        const kept = keptCell(entry.cell, next);
        if (kept.isErr()) return Result.err(kept.error);
        cells.push(kept.value);
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
            reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
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
          const kept = keptCell(entry.cell, next);
          if (kept.isErr()) return Result.err(kept.error);
          cells.push(kept.value);
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
      const added = newCells(rowIndex);
      if (added.isErr()) return Result.err(added.error);
      cells.push(...added.value);
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
      columnWidths.splice(
        op.column,
        0,
        ...Array.from({ length: count }, () => op.width),
      );
    else columnWidths.splice(op.column, 1);
  }
  const next = { ...table, rows };
  if (columnWidths !== undefined) next.columnWidths = columnWidths;
  if (earlierGrid.value !== undefined)
    next.formatting = { ...table.formatting, gridChange: earlierGrid.value };
  const gridStamp = stamps.value[rowStamps];
  const propertyStamp = stamps.value[rowStamps + 1];
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

/**
 * The table without the rows a merge closes, each taken out as removing a
 * row does: a merge none of the rows left continues is no merge.
 */
const withoutRows = (table: Table, closing: ReadonlySet<number>): Table => {
  let rows = table.rows;
  for (const index of [...closing].toSorted((left, right) => right - left))
    rows = rowsWithoutRow({ ...table, rows }, index);
  return { ...table, rows };
};

type ClosesOtherRowsOptions = {
  op: Extract<TableEditOp, { type: "mergeCells" }>;
  table: Table;
  closing: ReadonlySet<number>;
};
/**
 * A direct merge's rows that close go into its first row, which then holds
 * their words. Under a pending insertion or deletion of its own, the first
 * row would take them with it when that change removes it.
 */
const closesOtherRows = ({
  op,
  table,
  closing,
}: ClosesOtherRowsOptions): Result<void, DocumentOpRefusal> => {
  const first = table.rows[op.top]?.structuralChange;
  if (
    first === undefined ||
    [...closing].every((index) =>
      structurallyEqual(table.rows[index]?.structuralChange, first),
    )
  )
    return Result.ok(undefined);
  return refuse({
    op,
    reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
    message:
      "Resolve the pending insertion or deletion of the merge's first row before merging rows it would take whole.",
  });
};

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

type MergeReviewsOptions = {
  op: Extract<TableEditOp, { type: "mergeCells" }>;
  selected: readonly (readonly GridCell[])[];
  verticalOnly: boolean;
};
/**
 * Whether a direct merge keeps the table's pending records. A merged cell is
 * under no pending insertion, deletion or merge of its own, and a cell the
 * merge takes in records no other span, merge or width than it has. The
 * merged rows' own insertions and deletions stay on their words
 * (`reviewedMergeContent`).
 */
const mergeReviews = ({
  op,
  selected,
  verticalOnly,
}: MergeReviewsOptions): Result<void, DocumentOpRefusal> => {
  const conflict = (message: string) =>
    refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
      message,
    });
  const cells = selected.flat();
  if (cells.some(({ cell }) => cell.structuralChange !== undefined))
    return conflict(
      "Resolve the pending insertions, deletions and merges of the merged cells before merging.",
    );
  // A vertical merge keeps every cell; another keeps only the first.
  const absorbed = verticalOnly ? [] : cells.slice(1);
  if (
    absorbed.some(
      ({ cell }) =>
        !statesOwnGrid(cell.propertyChanges, cell.formatting, CELL_GRID_FIELDS),
    )
  )
    return conflict(
      "Resolve the pending span and width changes of the merged cells before merging.",
    );
  return Result.ok(undefined);
};

/** A merged cell's blocks from one row, and that row's own pending change. */
type MergedWords = {
  blocks: readonly BlockContent[];
  change: TableStructuralChangeInfo | undefined;
};

/** One paragraph of the merged cell, and the record its row's change gives it. */
type MergedParagraph = {
  paragraph: Paragraph;
  /** Which decision removes it, with the row's change it records. */
  removal:
    | { kind: typeof WRAP_KINDS.INSERTION; stamp: RevisionStamp }
    | { kind: typeof WRAP_KINDS.DELETION; stamp: RevisionStamp }
    | undefined;
};

/** The row change each kind of record goes with. */
const ROW_CHANGE_TYPES = {
  [WRAP_KINDS.INSERTION]: "tableRowInsertion",
  [WRAP_KINDS.DELETION]: "tableRowDeletion",
} as const satisfies Record<WrapKind, TableStructuralChangeInfo["type"]>;

const MARK_KINDS = {
  [WRAP_KINDS.INSERTION]: "ins",
  [WRAP_KINDS.DELETION]: "del",
} as const satisfies Record<WrapKind, "ins" | "del">;

/**
 * The paragraph's content recorded as the kind of change: words a pending
 * insertion already holds stay as they are, as a deletion keeps the words a
 * pending deletion holds.
 */
const recordedContent = (
  paragraph: Paragraph,
  removal: NonNullable<MergedParagraph["removal"]>,
): Result<ParagraphContent[], DocumentOpRefusal["reason"]> => {
  const wrap = (items: readonly ParagraphContent[]) => {
    if (items.length === 0) return Result.ok([]);
    const wrapped = wrapTracked({
      items,
      from: { offset: 0, zeroWidthBefore: 0 },
      to: {
        offset: paragraphLength({ type: "paragraph", content: [...items] }),
        zeroWidthBefore: Number.MAX_SAFE_INTEGER,
      },
      kind: removal.kind,
      stamp: removal.stamp,
      ...(removal.kind === WRAP_KINDS.DELETION
        ? { commentAnchors: "stand" as const }
        : {}),
    });
    switch (wrapped.kind) {
      case "refused":
        return Result.err(wrapped.reason);
      case "unchanged":
        return Result.ok([...items]);
      case "wrapped":
        return Result.ok(wrapped.content);
      default: {
        const unreachable: never = wrapped;
        return unreachable;
      }
    }
  };
  if (removal.kind === WRAP_KINDS.DELETION) return wrap(paragraph.content);
  // An insertion is never nested in another: the stretches between pending
  // insertions take their own.
  const content: ParagraphContent[] = [];
  let stretch: ParagraphContent[] = [];
  for (const item of [...paragraph.content, undefined]) {
    if (item !== undefined && item.type !== "insertion") {
      stretch.push(item);
      continue;
    }
    const wrapped = wrap(stretch);
    if (wrapped.isErr()) return wrapped;
    content.push(...wrapped.value);
    stretch = [];
    if (item !== undefined) content.push(item);
  }
  return Result.ok(content);
};

type ReviewedMergeContentOptions = {
  document: Document;
  op: Extract<TableEditOp, { type: "mergeCells" }>;
  /** The merged rows' own pending changes, one for each row. */
  changes: readonly (TableStructuralChangeInfo | undefined)[];
  words: readonly MergedWords[];
};
/**
 * A direct merge's cell content, across rows with different pending
 * insertions and deletions. The words of a row pending insertion keep that
 * insertion as records of their own, and those of a row pending deletion
 * that deletion, under its author, date and suggestion: rejecting the rows'
 * changes removes the inserted rows' words and keeps the deleted rows' in
 * place, and accepting them does the reverse. A paragraph mark between such
 * words and the next ones records the change too, so the words go with their
 * paragraphs. When the merged cell ends with words one decision removes,
 * the last paragraph that stays takes the cell's final mark as a paragraph
 * insertion or deletion at a container's end does.
 */
const reviewedMergeContent = ({
  document,
  op,
  changes,
  words,
}: ReviewedMergeContentOptions): Result<BlockContent[], DocumentOpRefusal> => {
  // Rows that stand or go together give their words as they are.
  if (changes.every((change) => structurallyEqual(change, changes.at(0))))
    return Result.ok(words.flatMap(({ blocks }) => blocks));
  const conflict = (message: string) =>
    refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
      message,
    });
  const merged: MergedParagraph[] = [];
  for (const { blocks: held, change } of words) {
    for (const block of held) {
      if (change === undefined && block.type === "paragraph") {
        merged.push({ paragraph: block, removal: undefined });
        continue;
      }
      if (
        block.type !== "paragraph" ||
        (change?.type !== "tableRowInsertion" &&
          change?.type !== "tableRowDeletion")
      )
        return conflict(
          "Resolve the pending row changes the merge crosses before merging cells that hold tables or content controls.",
        );
      // The records the words take are the row's change: its author, date
      // and suggestion, each under an id of its own.
      const { id, author, date, initials, suggestionId } = change.info;
      if (date === undefined)
        return conflict(
          "Resolve the undated pending row changes the merge crosses before merging.",
        );
      const stamp: RevisionStamp = {
        id,
        author,
        date,
        ...(initials === undefined ? {} : { initials }),
        ...(suggestionId === undefined ? {} : { suggestionId }),
      };
      merged.push({
        paragraph: block,
        removal:
          change.type === "tableRowInsertion"
            ? { kind: WRAP_KINDS.INSERTION, stamp }
            : { kind: WRAP_KINDS.DELETION, stamp },
      });
    }
  }
  const lastStanding = (kind: WrapKind) =>
    merged.findLastIndex(({ removal }) => removal?.kind !== kind);
  // Where a decision keeps a merged row, the merged cell keeps words of a
  // row it keeps; where it keeps none, the cell goes with them.
  const rowStands = (kind: WrapKind) =>
    changes.some((change) => change?.type !== ROW_CHANGE_TYPES[kind]);
  if (
    [WRAP_KINDS.INSERTION, WRAP_KINDS.DELETION].some(
      (kind) => rowStands(kind) && lastStanding(kind) === -1,
    )
  )
    return conflict(
      "Resolve the pending row changes the merge crosses before merging cells whose words would all go with them.",
    );
  const final = merged.length - 1;
  const after: Paragraph[] = [];
  for (const [index, { paragraph, removal }] of merged.entries()) {
    if (removal === undefined) {
      after.push(paragraph);
      continue;
    }
    const recorded = recordedContent(paragraph, removal);
    if (recorded.isErr())
      return refuse({
        op,
        reason: recorded.error,
        message: "The merged rows' words cannot record their pending change.",
      });
    const mark = MARK_KINDS[removal.kind];
    const existing = paragraph.pPrMark;
    if (
      existing !== undefined &&
      (existing.kind !== mark || existing.deletion !== undefined)
    )
      return conflict(
        "Resolve the pending paragraph mark changes of the merged rows before merging.",
      );
    // The cell's final mark stays; the paragraph that stays before it takes
    // the change below.
    const next: Paragraph = { ...paragraph, content: recorded.value };
    if (index === final) delete next.pPrMark;
    else if (existing === undefined)
      next.pPrMark = { kind: mark, info: stampInfo(removal.stamp) };
    after.push(next);
  }
  const ending = merged.at(final)?.removal;
  const standing = ending === undefined ? -1 : lastStanding(ending.kind);
  if (ending !== undefined && standing !== -1) {
    const survivor = after[standing] ?? panic("A merged paragraph is missing.");
    const last = after[final] ?? panic("A merged paragraph is missing.");
    if (
      survivor.pPrMark !== undefined ||
      survivor.propertyChanges !== undefined ||
      last.propertyChanges !== undefined
    )
      return conflict(
        "Resolve the pending paragraph changes at the end of the merged cell before merging.",
      );
    const info = stampInfo(ending.stamp);
    if (ending.kind === WRAP_KINDS.INSERTION) {
      // Rejected, each inserted break joins into the paragraph before it,
      // which ends up being the one that stood, with its properties.
      after[standing] = {
        ...survivor,
        pPrMark: { kind: "ins", info, survivor: "first" },
      };
      for (let index = standing + 1; index < final; index += 1) {
        const paragraph =
          after[index] ?? panic("A merged paragraph is missing.");
        if (paragraph.pPrMark?.kind === "ins")
          after[index] = {
            ...paragraph,
            pPrMark: { ...paragraph.pPrMark, survivor: "first" },
          };
      }
      if (
        !structurallyEqual(
          finalizeParagraphFormatting(survivor.formatting) ?? {},
          finalizeParagraphFormatting(last.formatting) ?? {},
        )
      )
        after[final] = {
          ...last,
          propertyChanges: [
            {
              type: "paragraphPropertyChange",
              info,
              ...(survivor.formatting === undefined
                ? {}
                : { previousFormatting: survivor.formatting }),
            },
          ],
        };
    } else {
      // Accepted, the last paragraph takes the words and properties of the
      // one that stood, as deleting the paragraphs after it does.
      after[standing] = { ...survivor, pPrMark: { kind: "del", info } };
      const formatting = finalizeParagraphFormatting(survivor.formatting);
      if (
        !structurallyEqual(
          formatting ?? {},
          finalizeParagraphFormatting(last.formatting) ?? {},
        )
      ) {
        const review = reviewFieldsOf(last);
        delete review.formatting;
        if (formatting !== undefined) review.formatting = formatting;
        review.propertyChanges = [
          foldedParagraphPropertyChange({
            paragraph: last,
            formatting,
            stamp: ending.stamp,
          }),
        ];
        after[final] = withReviewFields(last, review);
      }
    }
  }
  // The new records take fresh ids; those the rows' words held keep theirs.
  const before = merged.map(({ paragraph }) => paragraph);
  const held = packageIdentityKeyIndex(document.package);
  const own = new Map<string, number>();
  for (const key of identityKeysOf(before))
    own.set(key, (own.get(key) ?? 0) + 1);
  const fresh = freshenIdentities({
    before,
    after,
    newIds: op.newIds ?? {},
    usedElsewhere: () => ({
      has: (key) => held.count(key) > (own.get(key) ?? 0),
    }),
  });
  switch (fresh.kind) {
    case "needsIds":
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        message: `The merged rows' words need ${fresh.missing} more revision ids.`,
      });
    case "invalidId":
      return refuse({
        op,
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        message: `${fresh.id} cannot be a new revision id.`,
      });
    case "fresh":
      return Result.ok(fresh.paragraphs);
    default: {
      const unreachable: never = fresh;
      return unreachable;
    }
  }
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
  if (hasWrappedCellMarkup(table))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "Captures inside cell wrappers need a shared destination.",
    });
  if (
    op.revision !== undefined &&
    hasReviewBlockingMerge({ table, top: op.top, bottom: op.bottom })
  )
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
      message: "Resolve table reviews before changing cells.",
    });
  const selection = mergeSelection({ table, grid, op });
  if (selection.isErr()) return Result.err(selection.error);
  const { selected, verticalOnly } = selection.value;
  if (op.revision === undefined) {
    const kept = mergeReviews({ op, selected, verticalOnly });
    if (kept.isErr()) return Result.err(kept.error);
  }
  const earlierGrid = () =>
    new DocumentOpRefusal({
      opType: op.type,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
      message: "A pending change records another span than the merge changes.",
    });
  if (verticalOnly) {
    const fresh = freshCells({ document, op, ids: op.newBlockIds, count: 0 });
    if (fresh.isErr()) return Result.err(fresh.error);
    const stamps = stampsFor(op, selected.length * 2);
    if (stamps.isErr()) return Result.err(stamps.error);
    const rows: TableRow[] = [];
    for (const [rowIndex, row] of table.rows.entries()) {
      const entry = selected[rowIndex - op.top]?.at(0);
      if (entry === undefined) {
        rows.push(row);
        continue;
      }
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
      const kept =
        op.revision === undefined
          ? restatedCell(entry.cell, next, CELL_GRID_FIELDS).mapError(
              earlierGrid,
            )
          : Result.ok(next);
      if (kept.isErr()) return Result.err(kept.error);
      const cells = [...row.cells];
      cells[entry.index] = kept.value;
      rows.push({ ...row, cells });
    }
    const closing = closingRows({ before: table.rows, after: rows });
    if (closing.size === 0) return Result.ok({ ...table, rows });
    if (op.revision === undefined)
      return closesOtherRows({ op, table, closing }).map(() =>
        withoutRows({ ...table, rows }, closing),
      );
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
  const changes = table.rows
    .slice(op.top, op.bottom)
    .map(({ structuralChange }) => structuralChange);
  const holds = selected.some((row) =>
    row.some(({ cell }) => !isEmptyCell(cell)),
  );
  const words = selected.flatMap((row, offset) =>
    row
      .filter(({ cell }) => (holds ? !isEmptyCell(cell) : cell === first.cell))
      .map(({ cell }) => ({
        blocks: cell.content,
        change: changes[offset],
      })),
  );
  const reviewed = reviewedMergeContent({ document, op, changes, words });
  if (reviewed.isErr()) return Result.err(reviewed.error);
  const content = reviewed.value;
  const mergedCellWidth = mergedWidth(
    selected.at(0)?.map(({ cell }) => cell) ?? [],
  );
  if (mergedCellWidth.isErr())
    return mismatch(
      op,
      "Selected cell widths must be nonnegative safe integer units.",
    );
  const width = mergedCellWidth.value;
  const rows: TableRow[] = [];
  for (const [index, row] of table.rows.entries()) {
    const entries = selected[index - op.top];
    const leading = entries?.at(0);
    const base =
      index === op.top
        ? { ...first.cell, content }
        : fresh.value[index - op.top - 1];
    if (entries === undefined || leading === undefined || base === undefined) {
      rows.push(row);
      continue;
    }
    const cell = withSpan(base, op.right - op.left);
    const formatting = authored({ ...cell.formatting });
    if (index === op.top) {
      if (width === undefined) delete formatting.width;
      else formatting.width = width;
    }
    if (op.bottom - op.top > 1)
      formatting.vMerge = index === op.top ? "restart" : "continue";
    else delete formatting.vMerge;
    // The merged cell keeps the first cell's pending property changes; the
    // cells it takes in go with their own.
    const kept =
      index === op.top
        ? restatedCell(
            first.cell,
            { ...cell, formatting },
            CELL_GRID_FIELDS,
          ).mapError(earlierGrid)
        : Result.ok({ ...cell, formatting });
    if (kept.isErr()) return Result.err(kept.error);
    const cells = [...row.cells];
    cells.splice(leading.index, entries.length, kept.value);
    rows.push({ ...row, cells });
  }
  const closing = closingRows({ before: table.rows, after: rows });
  return closesOtherRows({ op, table, closing }).map(() =>
    withoutRows({ ...table, rows }, closing),
  );
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
  if (hasWrappedCellMarkup(table))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      message: "Captures inside cell wrappers need a shared destination.",
    });
  if (op.revision !== undefined && hasReview(table))
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
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
  // A direct split keeps the pending records beside it; the split cell is
  // under no pending insertion, deletion or merge of its own.
  if (
    op.revision === undefined &&
    group.some(({ cell }) => cell.structuralChange !== undefined)
  )
    return refuse({
      op,
      reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
      message:
        "Resolve the pending insertions, deletions and merges of the split cell before splitting it.",
    });
  // The new cells of a span take the cell's properties, and with them its
  // pending property changes, each under a record of its own.
  const copiedChanges =
    op.revision === undefined
      ? directRevisionIds({
          document,
          op,
          count: group.reduce(
            (sum, { cell }) =>
              sum + (cell.propertyChanges?.length ?? 0) * (span - 1),
            0,
          ),
        })
      : Result.ok([]);
  if (copiedChanges.isErr()) return Result.err(copiedChanges.error);
  const copiedIds = copiedChanges.value.values();
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
  const rows: TableRow[] = [];
  for (const [index, row] of table.rows.entries()) {
    const entry = group.find((cell) => cell.row === index);
    if (!entry) {
      rows.push(row);
      continue;
    }
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
    const kept =
      op.revision === undefined
        ? restatedCell(entry.cell, original, CELL_GRID_FIELDS).mapError(
            () =>
              new DocumentOpRefusal({
                opType: op.type,
                reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
                message:
                  "A pending change records another span than the split changes.",
              }),
          )
        : Result.ok(original);
    if (kept.isErr()) return Result.err(kept.error);
    const replacement: TableCell[] = [kept.value];
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
        const copied = copiedPropertyChanges({
          source: entry.cell,
          cell: next,
          ids: copiedIds,
        });
        if (copied.isErr())
          return refuse({
            op,
            reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
            message:
              "A pending change records another span than the split changes.",
          });
        replacement.push(copied.value);
      }
    }
    cells.splice(entry.index, 1, ...replacement);
    rows.push({ ...row, cells });
  }
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
    case DOCUMENT_OP_TYPES.INSERT_COLUMN: {
      const plans = recordsPlaceColumnsApart(table, grid.value)
        ? columnPlanAcrossRecords({
            table,
            grid: grid.value,
            op: { ...op, newBlockIds: [] },
          })
        : undefined;
      if (plans !== undefined)
        return Result.ok(
          plans.filter(
            (plan) => plan.type === "cells" && plan.added !== undefined,
          ).length * (op.count ?? 1),
        );
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
        }).length * (op.count ?? 1),
      );
    }
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
          reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
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
          reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
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
          reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
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
      const unmeasured = op.columnWidths.length === 0;
      if (
        (!unmeasured && op.columnWidths.length !== grid.value.width) ||
        op.columnWidths.some(
          (width) => !Number.isSafeInteger(width) || width <= 0,
        )
      )
        return mismatch(
          op,
          "Column widths must match the grid and be positive twip counts, or be none.",
        );
      if (table.formatting?.gridChange !== undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
          message: "Resolve the existing grid revision before resizing.",
        });
      if (table.propertyChanges !== undefined)
        return refuse({
          op,
          reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
          message:
            "Resolve the table property history before resizing its grid.",
        });
      const stamps = stampsFor(op, 2);
      if (stamps.isErr()) return Result.err(stamps.error);
      const { columnWidths: _previous, ...rest } = table;
      const next: Table = unmeasured
        ? rest
        : { ...rest, columnWidths: [...op.columnWidths] };
      const gridStamp = stamps.value.at(0);
      const propertyStamp = stamps.value.at(1);
      if (gridStamp !== undefined && propertyStamp !== undefined) {
        next.formatting = {
          ...table.formatting,
          gridChange: {
            id: gridStamp.id,
            columnWidths:
              table.columnWidths === undefined
                ? Array.from({ length: grid.value.width }, () => undefined)
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
