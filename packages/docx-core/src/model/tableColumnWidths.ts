import type { Table } from "./content";
import type { WritableRecord } from "./deepReadonly";
import { gridColumnsInRows, MAX_TABLE_GRID_COLUMNS } from "./tableGridLimits";

export const inferredTableColumnCount = ({
  rows,
}: Pick<Table, "rows">): number =>
  Math.min(gridColumnsInRows(rows), MAX_TABLE_GRID_COLUMNS);

export const declaredUnmeasuredColumnCount = (
  unmeasuredColumnCount: number | undefined,
): number | undefined =>
  unmeasuredColumnCount !== undefined &&
  Number.isSafeInteger(unmeasuredColumnCount) &&
  unmeasuredColumnCount > 0 &&
  unmeasuredColumnCount <= MAX_TABLE_GRID_COLUMNS
    ? unmeasuredColumnCount
    : undefined;

type TableColumnWidths = WritableRecord &
  Pick<
    Table,
    "columnWidths" | "omittedColumnWidthIndices" | "unmeasuredColumnCount"
  >;

/** Project the grid's declared widths onto layout's numeric array. */
export const tableColumnWidthFields = (
  stated: readonly (number | undefined)[],
  inferColumnCount?: () => number,
): TableColumnWidths => {
  const columnWidths: number[] = [];
  const omittedColumnWidthIndices: number[] = [];
  for (const [index, width] of stated.entries()) {
    columnWidths.push(width ?? 0);
    if (width === undefined) omittedColumnWidthIndices.push(index);
  }
  if (omittedColumnWidthIndices.length === columnWidths.length) {
    return inferColumnCount !== undefined &&
      columnWidths.length > 0 &&
      columnWidths.length !== inferColumnCount()
      ? { unmeasuredColumnCount: columnWidths.length }
      : {};
  }
  return {
    columnWidths,
    ...(omittedColumnWidthIndices.length === 0
      ? {}
      : { omittedColumnWidthIndices }),
  };
};

/** Recover the declared widths, including omitted attributes, from a live grid. */
export const statedTableColumnWidths = ({
  columnWidths,
  omittedColumnWidthIndices,
  unmeasuredColumnCount,
}: {
  columnWidths?: readonly number[];
  omittedColumnWidthIndices?: readonly number[];
  unmeasuredColumnCount?: number;
}): (number | undefined)[] | undefined => {
  if (columnWidths === undefined) {
    const count = declaredUnmeasuredColumnCount(unmeasuredColumnCount);
    return count === undefined
      ? undefined
      : Array.from({ length: count }, () => undefined);
  }
  const omitted = new Set(omittedColumnWidthIndices);
  return columnWidths.map((width, index) =>
    width === 0 && omitted.has(index) ? undefined : width,
  );
};

export const withTableColumnWidths = (
  table: Table,
  stated: readonly (number | undefined)[],
): Table => {
  const {
    columnWidths: _widths,
    omittedColumnWidthIndices: _omitted,
    unmeasuredColumnCount: _unmeasured,
    ...rest
  } = table;
  return {
    ...rest,
    ...tableColumnWidthFields(stated, () => inferredTableColumnCount(table)),
  };
};
