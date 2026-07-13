import type { TableCell, TableRow } from "../types";

export type TableCellGrid = {
  occupiedColumnsByRow: ReadonlyMap<number, ReadonlySet<number>>;
  sourceCellsByRow: ReadonlyMap<number, ReadonlyMap<number, TableCell>>;
  sourceColumnsByCell: ReadonlyMap<TableCell, number>;
};

export const buildTableCellGrid = (
  rows: readonly TableRow[],
  columnCount: number,
): TableCellGrid => {
  const occupiedColumnsByRow = new Map<number, Set<number>>();
  const sourceCellsByRow = new Map<number, Map<number, TableCell>>();
  const sourceColumnsByCell = new Map<TableCell, number>();

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    if (!row) {
      continue;
    }

    let columnIndex = firstAvailableColumn(occupiedColumnsByRow.get(rowIndex), row.gridBefore ?? 0);
    for (const cell of row.cells) {
      const columnSpan = cell.colSpan ?? 1;
      const rowSpan = cell.rowSpan ?? 1;
      sourceColumnsByCell.set(cell, columnIndex);

      const rowEnd = Math.min(rows.length, rowIndex + rowSpan);
      const columnEnd = Math.min(columnCount, columnIndex + columnSpan);
      for (let gridRowIndex = rowIndex; gridRowIndex < rowEnd; gridRowIndex++) {
        const cellsByColumn = getOrCreateCellRow(sourceCellsByRow, gridRowIndex);
        for (let gridColumnIndex = columnIndex; gridColumnIndex < columnEnd; gridColumnIndex++) {
          cellsByColumn.set(gridColumnIndex, cell);
        }
      }

      if (rowSpan > 1) {
        for (
          let spannedRowIndex = rowIndex + 1;
          spannedRowIndex < rowIndex + rowSpan;
          spannedRowIndex++
        ) {
          const occupied = getOrCreateOccupiedRow(occupiedColumnsByRow, spannedRowIndex);
          for (let columnOffset = 0; columnOffset < columnSpan; columnOffset++) {
            occupied.add(columnIndex + columnOffset);
          }
        }
      }

      columnIndex = firstAvailableColumn(
        occupiedColumnsByRow.get(rowIndex),
        columnIndex + columnSpan,
      );
    }
  }

  return { occupiedColumnsByRow, sourceCellsByRow, sourceColumnsByCell };
};

export const getFirstAvailableColumn = (
  grid: TableCellGrid,
  rowIndex: number,
  startingColumn: number,
): number => firstAvailableColumn(grid.occupiedColumnsByRow.get(rowIndex), startingColumn);

export const getSourceCellAt = (
  grid: TableCellGrid,
  rowIndex: number,
  columnIndex: number,
): TableCell | undefined => grid.sourceCellsByRow.get(rowIndex)?.get(columnIndex);

export const getSourceCellColumn = (grid: TableCellGrid, cell: TableCell): number | undefined =>
  grid.sourceColumnsByCell.get(cell);

export const getTableCellVerticalBorderHeight = (
  grid: TableCellGrid,
  cell: TableCell | undefined,
  rowIndex: number,
): number => {
  const sourceColumn = cell ? getSourceCellColumn(grid, cell) : undefined;
  const aboveCell =
    sourceColumn === undefined ? undefined : getSourceCellAt(grid, rowIndex - 1, sourceColumn);
  const aboveBottom = aboveCell?.borders?.bottom;
  const aboveOwnsEdge =
    aboveBottom !== undefined &&
    aboveBottom.width !== 0 &&
    aboveBottom.style !== "none" &&
    aboveBottom.style !== "nil";
  const top = rowIndex === 0 || !aboveOwnsEdge ? (cell?.borders?.top?.width ?? 0) : 0;
  const bottom = cell?.borders?.bottom?.width ?? 0;
  return top + bottom;
};

const firstAvailableColumn = (
  occupiedColumns: ReadonlySet<number> | undefined,
  startingColumn: number,
): number => {
  let columnIndex = startingColumn;
  while (occupiedColumns?.has(columnIndex)) {
    columnIndex++;
  }
  return columnIndex;
};

const getOrCreateCellRow = (
  sourceCellsByRow: Map<number, Map<number, TableCell>>,
  rowIndex: number,
): Map<number, TableCell> => {
  const existing = sourceCellsByRow.get(rowIndex);
  if (existing !== undefined) {
    return existing;
  }
  const cellsByColumn = new Map<number, TableCell>();
  sourceCellsByRow.set(rowIndex, cellsByColumn);
  return cellsByColumn;
};

const getOrCreateOccupiedRow = (
  occupiedColumnsByRow: Map<number, Set<number>>,
  rowIndex: number,
): Set<number> => {
  const existing = occupiedColumnsByRow.get(rowIndex);
  if (existing !== undefined) {
    return existing;
  }
  const occupiedColumns = new Set<number>();
  occupiedColumnsByRow.set(rowIndex, occupiedColumns);
  return occupiedColumns;
};
