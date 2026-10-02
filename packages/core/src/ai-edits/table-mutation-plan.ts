import type { Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";

export type TableRectangle = {
  left: number;
  top: number;
  right: number;
  bottom: number;
};

export type TableMutationPlanTarget =
  | { type: "none" }
  | { type: "tableStructure"; tablePosition: number }
  | {
      type: "mergeCells";
      tablePosition: number;
      rectangle: TableRectangle;
      foldedRows: ReturnType<typeof tableMergeFoldedRows>;
    }
  | { type: "splitCell"; tablePosition: number; rectangle: TableRectangle };

type TableMutationPlanCandidate<T> = {
  item: T;
  operationId: string;
  target: TableMutationPlanTarget;
};

type TableMutationPlanSkip = {
  id: string;
  reason: "noopOperation" | "unsupportedBlock";
};

type TableMutationPlan<T> = {
  executable: T[];
  skipped: TableMutationPlanSkip[];
};

export const planTableMutations = <T>(
  candidates: readonly TableMutationPlanCandidate<T>[],
): TableMutationPlan<T> => {
  const tableStructureMutations = new Set<number>();
  const mergeTables = new Set<number>();
  const splitTables = new Set<number>();

  for (const { target } of candidates) {
    switch (target.type) {
      case "tableStructure":
        tableStructureMutations.add(target.tablePosition);
        break;
      case "mergeCells":
        mergeTables.add(target.tablePosition);
        break;
      case "splitCell":
        splitTables.add(target.tablePosition);
        break;
      case "none":
        break;
    }
  }

  const mergeRectanglesByTable = new Map<number, TableRectangle[]>();
  const foldedRowsByTable = new Map<number, Map<number, number>>();
  const splitRectanglesByTable = new Map<number, TableRectangle[]>();
  const executable: T[] = [];
  const skipped: TableMutationPlanSkip[] = [];

  for (const candidate of candidates) {
    const { target } = candidate;
    if (target.type === "mergeCells") {
      if (
        tableStructureMutations.has(target.tablePosition) ||
        splitTables.has(target.tablePosition)
      ) {
        skipped.push({ id: candidate.operationId, reason: "unsupportedBlock" });
        continue;
      }
      const claimedRectangles = mergeRectanglesByTable.get(target.tablePosition) ?? [];
      if (
        claimedRectangles.some((rectangle) => tableRectanglesEqual(rectangle, target.rectangle))
      ) {
        skipped.push({ id: candidate.operationId, reason: "noopOperation" });
        continue;
      }
      if (
        claimedRectangles.some((rectangle) => tableRectanglesOverlap(rectangle, target.rectangle))
      ) {
        skipped.push({ id: candidate.operationId, reason: "unsupportedBlock" });
        continue;
      }
      const foldedRows = foldedRowsByTable.get(target.tablePosition) ?? new Map<number, number>();
      // Disjoint rectangles can still depend on the same row. Refuse the
      // later merge when together they remove its last starting cell: direct
      // mode closes that row, but tracked mode cannot record its closure.
      if (
        target.foldedRows.some(({ row, cellCount, foldedCount }) => {
          const claimedCount = foldedRows.get(row) ?? 0;
          return claimedCount > 0 && claimedCount + foldedCount >= cellCount;
        })
      ) {
        skipped.push({ id: candidate.operationId, reason: "unsupportedBlock" });
        continue;
      }
      for (const { row, foldedCount } of target.foldedRows) {
        foldedRows.set(row, (foldedRows.get(row) ?? 0) + foldedCount);
      }
      foldedRowsByTable.set(target.tablePosition, foldedRows);
      claimedRectangles.push(target.rectangle);
      mergeRectanglesByTable.set(target.tablePosition, claimedRectangles);
      executable.push(candidate.item);
      continue;
    }

    if (target.type === "splitCell") {
      if (
        tableStructureMutations.has(target.tablePosition) ||
        mergeTables.has(target.tablePosition)
      ) {
        skipped.push({ id: candidate.operationId, reason: "unsupportedBlock" });
        continue;
      }
      const claimedRectangles = splitRectanglesByTable.get(target.tablePosition) ?? [];
      if (
        claimedRectangles.some((rectangle) => tableRectanglesEqual(rectangle, target.rectangle))
      ) {
        skipped.push({ id: candidate.operationId, reason: "noopOperation" });
        continue;
      }
      if (
        claimedRectangles.some((rectangle) => tableRectanglesOverlap(rectangle, target.rectangle))
      ) {
        skipped.push({ id: candidate.operationId, reason: "unsupportedBlock" });
        continue;
      }
      claimedRectangles.push(target.rectangle);
      splitRectanglesByTable.set(target.tablePosition, claimedRectangles);
      executable.push(candidate.item);
      continue;
    }

    executable.push(candidate.item);
  }

  return { executable, skipped };
};

/** Count physical cells folded away, excluding the retained origin cell. */
export const tableMergeFoldedRows = (table: PMNode, rectangle: TableRectangle) => {
  const map = TableMap.get(table);
  const rows = new Map<number, { row: number; cellCount: number; foldedCount: number }>();
  for (const position of map.cellsInRect(rectangle).slice(1)) {
    const row = map.findCell(position).top;
    const counts = rows.get(row) ?? { row, cellCount: table.child(row).childCount, foldedCount: 0 };
    counts.foldedCount++;
    rows.set(row, counts);
  }
  return [...rows.values()];
};

const tableRectanglesOverlap = (left: TableRectangle, right: TableRectangle): boolean =>
  left.left < right.right &&
  right.left < left.right &&
  left.top < right.bottom &&
  right.top < left.bottom;

export const tableRectanglesEqual = (left: TableRectangle, right: TableRectangle): boolean =>
  left.left === right.left &&
  left.top === right.top &&
  left.right === right.right &&
  left.bottom === right.bottom;
