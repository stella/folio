import { Result } from "better-result";

import type { TableCell, TableRow } from "../model/document";
import { structurallyEqual } from "./equality";

/** The grid fields a cell merge or split rewrites, each with its unstated value. */
export const CELL_GRID_FIELDS = {
  gridSpan: 1,
  vMerge: undefined,
  width: undefined,
} as const;
/** The cell field a column edit rewrites. */
export const CELL_SPAN_FIELDS = { gridSpan: 1 } as const;
/** The row fields a column edit rewrites. */
export const ROW_GRID_FIELDS = { gridBefore: 0, gridAfter: 0 } as const;
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
export const statesOwnGrid = (
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
export const restatedCell = (
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

/** `after`, the direct edit of `before`, with `before`'s pending property changes restated. */
export const restatedRow = (
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
