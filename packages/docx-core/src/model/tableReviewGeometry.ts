import type { Table, TableCell, TableRow } from "./content";
import { withTableColumnWidths } from "./tableColumnWidths";

/** Physical cell membership and formatting while geometry review is pending. */
type TableReviewGeometry = {
  readonly columnWidths: readonly (number | undefined)[];
  readonly holdsCell: (cell: TableCell) => boolean;
  readonly rowFormatting: (row: TableRow) => TableRow["formatting"];
  readonly cellFormatting: (cell: TableCell) => TableCell["formatting"];
};

/** The physical geometry while the deleted cells still occupy their slots. */
export const tableReviewGeometry = (
  table: Table,
): TableReviewGeometry | undefined => {
  const grid = table.formatting?.gridChange;
  if (grid === undefined || !Array.isArray(grid.geometryChange?.revisionIds))
    return undefined;
  const ids = new Set(grid.geometryChange.revisionIds);
  return {
    columnWidths: grid.columnWidths,
    holdsCell: (cell: TableCell) => {
      const change = cell.structuralChange;
      return change?.type !== "tableCellInsertion" || !ids.has(change.info.id);
    },
    rowFormatting: (row: TableRow) => {
      const prior = row.propertyChanges?.find(({ info }) => ids.has(info.id));
      return prior === undefined ? row.formatting : prior.previousFormatting;
    },
    cellFormatting: (cell: TableCell) => {
      const prior = cell.propertyChanges?.find(({ info }) => ids.has(info.id));
      return prior === undefined ? cell.formatting : prior.previousFormatting;
    },
  };
};

/** A view of the pending rows with their physical spans and omissions. */
export const tableWithPhysicalColumns = (table: Table): Table => {
  const geometry = tableReviewGeometry(table);
  if (geometry === undefined) return table;
  const widths = geometry.columnWidths;
  const rows: TableRow[] = [];
  for (const row of table.rows) {
    const cells: TableCell[] = [];
    const offsets = [0];
    for (const [index, cell] of row.cells.entries()) {
      if (!geometry.holdsCell(cell)) {
        offsets[index + 1] = cells.length;
        continue;
      }
      const formatting = geometry.cellFormatting(cell);
      const next = { ...cell };
      if (formatting === undefined) delete next.formatting;
      else next.formatting = formatting;
      cells.push(next);
      offsets[index + 1] = cells.length;
    }
    const formatting = geometry.rowFormatting(row);
    const next = { ...row, cells };
    if (formatting === undefined) delete next.formatting;
    else next.formatting = formatting;
    if (cells.length !== row.cells.length) {
      const relocate = <T extends { index: number }>(items: T[]) =>
        items.map((item) =>
          Object.assign({}, item, {
            index: offsets[item.index] ?? cells.length,
          }),
        );
      if (row.preserved?.children !== undefined)
        next.preserved = {
          ...row.preserved,
          children: relocate(row.preserved.children),
        };
      if (row.bookmarks !== undefined) next.bookmarks = relocate(row.bookmarks);
    }
    rows.push(next);
  }
  return withTableColumnWidths({ ...table, rows }, widths);
};
