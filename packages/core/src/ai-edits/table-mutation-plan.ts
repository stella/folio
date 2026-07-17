export type TableRectangle = {
  left: number;
  top: number;
  right: number;
  bottom: number;
};

export type TableMutationPlanTarget =
  | { type: "none" }
  | { type: "tableStructure"; tablePosition: number }
  | { type: "mergeCells"; tablePosition: number; rectangle: TableRectangle }
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
    if (target.type === "tableStructure") {
      tableStructureMutations.add(target.tablePosition);
      continue;
    }
    if (target.type === "mergeCells") {
      mergeTables.add(target.tablePosition);
      continue;
    }
    if (target.type === "splitCell") {
      splitTables.add(target.tablePosition);
    }
  }

  const mergeRectanglesByTable = new Map<number, TableRectangle[]>();
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
    }

    executable.push(candidate.item);
  }

  return { executable, skipped };
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
