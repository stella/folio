import { panic } from "better-result";
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
  doc: PMNode,
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
      if (
        claimedRectangles.length > 0 &&
        mergesRemoveRow({
          doc,
          tablePosition: target.tablePosition,
          rectangles: [...claimedRectangles, target.rectangle],
        })
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
      executable.push(candidate.item);
      continue;
    }

    executable.push(candidate.item);
  }

  return { executable, skipped };
};

type MergesRemoveRowOptions = {
  doc: PMNode;
  tablePosition: number;
  rectangles: readonly TableRectangle[];
};

/**
 * Disjoint merges can still share a row's survival. Direct merging closes a
 * row left without cells, invalidating other rectangles' original row indices;
 * tracked merging cannot record that removal. Keep a row-removing merge alone
 * and refuse the later merge when their combined folded cells empty a row.
 */
const mergesRemoveRow = ({ doc, tablePosition, rectangles }: MergesRemoveRowOptions): boolean => {
  const table = doc.nodeAt(tablePosition);
  if (table?.type.spec["tableRole"] !== "table") {
    return panic("A planned cell merge lost its table", { position: tablePosition });
  }
  const map = TableMap.get(table);
  const folded = new Set<number>();
  for (const rectangle of rectangles) {
    const kept = map.map[rectangle.top * map.width + rectangle.left];
    for (const position of map.cellsInRect(rectangle)) {
      if (position !== kept) folded.add(position);
    }
  }
  let removesRow = false;
  table.forEach((row, rowOffset) => {
    let remaining = row.childCount;
    row.forEach((_cell, cellOffset) => {
      if (folded.has(rowOffset + 1 + cellOffset)) remaining--;
    });
    if (row.childCount > 0 && remaining === 0) removesRow = true;
  });
  return removesRow;
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
