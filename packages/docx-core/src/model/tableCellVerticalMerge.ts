/** Read the modeled vertical-merge state; absent formatting names no merge. */
import type { TableCellFormatting } from "./document";

export const TABLE_CELL_VERTICAL_MERGES = {
  NONE: "none",
  RESTART: "restart",
  CONTINUE: "continue",
} as const;

export const tableCellVerticalMerge = (value: TableCellFormatting["vMerge"]) =>
  value ?? TABLE_CELL_VERTICAL_MERGES.NONE;
