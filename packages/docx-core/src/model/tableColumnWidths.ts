import type { Table } from "./content";
import type { WritableRecord } from "./deepReadonly";

type TableColumnWidths = WritableRecord &
  Pick<Table, "columnWidths" | "omittedColumnWidthIndices">;

/** Project the grid's declared widths onto layout's numeric array. */
export const tableColumnWidthFields = (
  stated: readonly (number | undefined)[],
): TableColumnWidths => {
  const columnWidths: number[] = [];
  const omittedColumnWidthIndices: number[] = [];
  for (const [index, width] of stated.entries()) {
    columnWidths.push(width ?? 0);
    if (width === undefined) omittedColumnWidthIndices.push(index);
  }
  if (omittedColumnWidthIndices.length === columnWidths.length) return {};
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
}: {
  columnWidths?: readonly number[];
  omittedColumnWidthIndices?: readonly number[];
}): (number | undefined)[] | undefined => {
  const omitted = new Set(omittedColumnWidthIndices);
  return columnWidths?.map((width, index) =>
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
    ...rest
  } = table;
  return { ...rest, ...tableColumnWidthFields(stated) };
};
