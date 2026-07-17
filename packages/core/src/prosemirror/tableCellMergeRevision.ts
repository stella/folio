import type { TableCell } from "../types/document";

export type TableCellMergeChange = Extract<
  NonNullable<TableCell["structuralChange"]>,
  { type: "tableCellMerge" }
>;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const isVerticalMergeRevisionValue = (
  value: unknown,
): value is NonNullable<TableCellMergeChange["verticalMerge"]> =>
  value === "continue" || value === "rest";

export const getTableCellMergeChange = (value: unknown): TableCellMergeChange | null => {
  if (!isRecord(value)) {
    return null;
  }
  const change = value["structuralChange"];
  if (!isRecord(change) || change["type"] !== "tableCellMerge") {
    return null;
  }
  const info = change["info"];
  if (!isRecord(info) || typeof info["id"] !== "number" || typeof info["author"] !== "string") {
    return null;
  }

  return {
    type: "tableCellMerge",
    info: {
      id: info["id"],
      author: info["author"],
      ...(typeof info["date"] === "string" ? { date: info["date"] } : {}),
    },
    ...(isVerticalMergeRevisionValue(change["verticalMerge"])
      ? { verticalMerge: change["verticalMerge"] }
      : {}),
    ...(isVerticalMergeRevisionValue(change["verticalMergeOriginal"])
      ? { verticalMergeOriginal: change["verticalMergeOriginal"] }
      : {}),
  };
};
