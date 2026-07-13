import type { TableRow } from "./types";

const ownsNestedGrid = (row: TableRow): boolean => {
  const cell = row.cells.length === 1 ? row.cells.at(0) : undefined;
  if (!cell?.blocks.some((block) => block.kind === "table")) {
    return false;
  }

  return cell.blocks.every(
    (block) => block.kind === "table" || (block.kind === "paragraph" && block.runs.length === 0),
  );
};

export const getTableRowLeadingWidth = (row: TableRow, columnWidths: readonly number[]): number => {
  if (ownsNestedGrid(row)) {
    return 0;
  }

  return columnWidths
    .slice(0, row.gridBefore ?? 0)
    .reduce((sum, columnWidth) => sum + columnWidth, 0);
};
