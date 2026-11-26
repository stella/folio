import type { Table, TableCell, TableRow } from "./content";

/** The physical geometry while the deleted cells still occupy their slots. */
export const columnDeletionReviewGeometry = (table: Table) => {
  const grid = table.formatting?.gridChange;
  if (grid === undefined || !Array.isArray(grid.columnDeletionRevisionIds))
    return undefined;
  const ids = new Set(grid.columnDeletionRevisionIds);
  return {
    columnWidths: grid.columnWidths,
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
  const geometry = columnDeletionReviewGeometry(table);
  if (geometry === undefined) return table;
  const widths = geometry.columnWidths;
  const rows: TableRow[] = [];
  for (const row of table.rows) {
    const cells: TableCell[] = [];
    for (const cell of row.cells) {
      const formatting = geometry.cellFormatting(cell);
      const next = { ...cell };
      if (formatting === undefined) delete next.formatting;
      else next.formatting = formatting;
      cells.push(next);
    }
    const formatting = geometry.rowFormatting(row);
    const next = { ...row, cells };
    if (formatting === undefined) delete next.formatting;
    else next.formatting = formatting;
    rows.push(next);
  }
  const pending = { ...table, rows };
  if (widths.every((width) => width !== undefined))
    pending.columnWidths = [...widths];
  else delete pending.columnWidths;
  return pending;
};
