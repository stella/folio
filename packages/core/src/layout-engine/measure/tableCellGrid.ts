import type { TableCell, TableRow } from "../types";

export type TableCellGrid = {
  occupiedColumnsByRow: ReadonlyMap<number, ReadonlySet<number>>;
  sourceCellsByRow: ReadonlyMap<number, ReadonlyMap<number, TableCell>>;
  sourceColumnsByCell: ReadonlyMap<TableCell, number>;
};

export type TableCellPlacement = {
  sourceColumn: number;
  columnSpan: number;
  left: number;
  width: number;
};

export type TableCellPlacements = ReadonlyMap<TableCell, TableCellPlacement>;

type BuildTableCellPlacementsOptions = {
  grid: TableCellGrid;
  columnWidths: readonly number[];
  bidi: boolean;
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

    const gridBefore = Math.max(0, Math.min(columnCount, Math.trunc(row.gridBefore ?? 0)));
    let columnIndex = firstAvailableColumn(occupiedColumnsByRow.get(rowIndex), gridBefore);
    for (const cell of row.cells) {
      const columnSpan = Math.max(1, Math.trunc(cell.colSpan ?? 1));
      const rowSpan = Math.max(1, Math.trunc(cell.rowSpan ?? 1));
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
        for (let spannedRowIndex = rowIndex + 1; spannedRowIndex < rowEnd; spannedRowIndex++) {
          const occupied = getOrCreateOccupiedRow(occupiedColumnsByRow, spannedRowIndex);
          for (let gridColumnIndex = columnIndex; gridColumnIndex < columnEnd; gridColumnIndex++) {
            occupied.add(gridColumnIndex);
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

/** Resolve all source cells to canonical logical columns and physical boxes. */
export const buildTableCellPlacements = ({
  grid,
  columnWidths,
  bidi,
}: BuildTableCellPlacementsOptions): TableCellPlacements => {
  const columnOffsets = [0];
  for (const columnWidth of columnWidths) {
    columnOffsets.push((columnOffsets.at(-1) ?? 0) + columnWidth);
  }
  const tableWidth = columnOffsets.at(-1) ?? 0;
  const placements = new Map<TableCell, TableCellPlacement>();

  for (const [cell, sourceColumn] of grid.sourceColumnsByCell) {
    if (sourceColumn < 0 || sourceColumn >= columnWidths.length) {
      continue;
    }
    const declaredColumnSpan = Math.max(1, Math.trunc(cell.colSpan ?? 1));
    const columnSpan = Math.min(declaredColumnSpan, columnWidths.length - sourceColumn);
    const logicalLeft = columnOffsets.at(sourceColumn) ?? 0;
    const logicalRight = columnOffsets.at(sourceColumn + columnSpan) ?? logicalLeft;
    const width = logicalRight - logicalLeft;
    placements.set(cell, {
      sourceColumn,
      columnSpan,
      left: bidi ? tableWidth - logicalRight : logicalLeft,
      width,
    });
  }

  return placements;
};

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
    aboveBottom !== undefined && aboveBottom.style !== "none" && aboveBottom.style !== "nil";
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
