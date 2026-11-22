import { readTableWidthKind } from "../../types/tableWidthKind";
import {
  resolveTableCellPadding,
  resolveTableWidthPx,
  type TableBlock,
  type TableCell,
  tableColumnsArePinned,
} from "../types";
import {
  getSourceCellColumn,
  getTableCellSpacingInsetsX,
  type TableCellGrid,
} from "./tableCellGrid";

export const isAutoWidthNoWrapCell = (cell: TableCell): boolean =>
  cell.noWrap === true &&
  readTableWidthKind(cell.preferredWidth?.type) === "automatic";

/** An authored dxa noWrap holds the width and lets the text wrap. */
export const tableCellKeepsNoWrap = (
  table: TableBlock,
  cell: TableCell,
): boolean =>
  cell.noWrap === true &&
  table.layout !== "fixed" &&
  readTableWidthKind(cell.preferredWidth?.type) !== "twips" &&
  (isAutoWidthNoWrapCell(cell) || !tableColumnsArePinned(table));

type CellWidthConstraint = {
  start: number;
  end: number;
  minimum: number;
  requiredAutoWidth: number;
};

type ExpandAutoWidthNoWrapColumnsOptions = {
  table: TableBlock;
  grid: TableCellGrid;
  columnWidths: number[];
  unwrappedContentWidth: (cell: TableCell, currentWidth: number) => number;
};

const widthOfRange = (columnWidths: number[], start: number, end: number) =>
  columnWidths.slice(start, end).reduce((total, width) => total + width, 0);

export const expandAutoWidthNoWrapColumns = ({
  table,
  grid,
  columnWidths,
  unwrappedContentWidth,
}: ExpandAutoWidthNoWrapColumnsOptions): number[] => {
  if (
    table.layout === "fixed" ||
    !table.rows.some(
      (row) => !row.hidden && row.cells.some(isAutoWidthNoWrapCell),
    )
  ) {
    return columnWidths;
  }

  // Constraints are over logical column intervals, including spans and row
  // omissions. Keep one bound per interval even when many rows share it.
  const constraints = new Map<number, CellWidthConstraint>();
  const pinnedColumns = new Set<number>();
  for (const row of table.rows) {
    if (row.hidden) continue;
    for (const cell of row.cells) {
      const start = getSourceCellColumn(grid, cell);
      if (start === undefined || start >= columnWidths.length) continue;
      const span = Math.max(1, cell.colSpan ?? 1);
      const end = Math.min(start + span, columnWidths.length);
      const currentWidth = widthOfRange(columnWidths, start, end);
      const { start: spacingStart, end: spacingEnd } =
        getTableCellSpacingInsetsX(
          table.cellSpacing,
          start,
          span,
          columnWidths.length,
        );
      const cellSpacing = spacingStart + spacingEnd;
      const { left, right } = resolveTableCellPadding(cell);
      const insets = cellSpacing + left + right;
      const autoNoWrap = isAutoWidthNoWrapCell(cell);
      const statedWidth = cell.preferredWidth;
      const pinnedWidth =
        cell.noWrap &&
        statedWidth &&
        readTableWidthKind(statedWidth.type) === "twips"
          ? resolveTableWidthPx(statedWidth.value, statedWidth.type, 0)
          : undefined;
      if (pinnedWidth !== undefined && end === start + 1) {
        pinnedColumns.add(start);
      }
      const contentWidth =
        pinnedWidth === undefined
          ? unwrappedContentWidth(cell, Math.max(1, currentWidth - insets)) +
            insets
          : pinnedWidth + cellSpacing;
      // Ordinary cells can give up unused width. Do not recompute the grid
      // for an ordinary cell whose content already wraps in its current slot.
      const minimum =
        autoNoWrap || pinnedWidth !== undefined
          ? contentWidth
          : Math.min(currentWidth, contentWidth);
      const key = start * columnWidths.length + end - 1;
      const constraint = constraints.get(key) ?? {
        start,
        end,
        minimum: 0,
        requiredAutoWidth: 0,
      };
      constraint.minimum = Math.max(constraint.minimum, minimum);
      if (autoNoWrap) {
        constraint.requiredAutoWidth = Math.max(
          constraint.requiredAutoWidth,
          contentWidth,
        );
      }
      constraints.set(key, constraint);
    }
  }

  const bounds = Array.from(constraints.values());
  if (
    !bounds.some(
      ({ start, end, requiredAutoWidth }) =>
        requiredAutoWidth > widthOfRange(columnWidths, start, end),
    )
  ) {
    return columnWidths;
  }
  const expanded = [...columnWidths];
  let addedWidth = 0;
  for (const { start, end, minimum } of bounds) {
    const extra = minimum - widthOfRange(expanded, start, end);
    if (extra <= 0) continue;
    const columns: number[] = [];
    for (let index = start; index < end; index++) {
      if (!pinnedColumns.has(index)) columns.push(index);
    }
    if (columns.length === 0) {
      for (let index = start; index < end; index++) columns.push(index);
    }
    const share = extra / columns.length;
    for (const index of columns) {
      expanded[index] = (expanded[index] ?? 0) + share;
    }
    addedWidth += extra;
  }

  // Reclaim unused room in the original grid before growing the table.
  // Track remaining slack on each interval as columns give room back.
  const intervals = bounds.map(({ start, end, minimum }) => ({
    start,
    end,
    slack: widthOfRange(expanded, start, end) - minimum,
  }));
  for (let index = 0; index < expanded.length && addedWidth > 0; index++) {
    const width = expanded[index] ?? 0;
    let reclaimed = Math.min(addedWidth, width - 1);
    const affected = intervals.filter(
      ({ start, end }) => start <= index && index < end,
    );
    for (const interval of affected) {
      reclaimed = Math.min(reclaimed, interval.slack);
    }
    if (reclaimed <= 0) continue;
    expanded[index] = width - reclaimed;
    addedWidth -= reclaimed;
    for (const interval of affected) {
      interval.slack -= reclaimed;
    }
  }
  return expanded;
};
