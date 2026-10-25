import {
  type BatchOrderAddress,
  contentRank,
  groupInsertionRuns,
  orderBatchOperations,
} from "@stll/docx-core/ops";
import type { FolioDocumentOperation } from "@stll/folio-core/server";

import type { AppliedOperation, Mode } from "./operations.ts";
import type { Row } from "./oracle.ts";

type AnyOperation = AppliedOperation;
type Handle = { blockId: string; startOffset: number; endOffset: number };

const INSERTIONS = new Set([
  "insertAfterBlock",
  "insertBeforeBlock",
  "insertTable",
]);
const ANNOTATIONS = new Set([
  "formatRange",
  "commentOnRange",
  "commentOnBlock",
]);
const RANGED = new Set(["replaceRange", "formatRange", "commentOnRange"]);

const handleOf = (operation: AnyOperation): Handle | null =>
  RANGED.has(operation.type) ? (operation["range"] as Handle) : null;
const blockIdOf = (operation: AnyOperation): string =>
  handleOf(operation)?.blockId ?? String(operation["blockId"]);

const placement = (
  operation: AnyOperation,
  preRows: readonly Row[],
): [number, number, number] => {
  const index = preRows.findIndex((row) => row.id === blockIdOf(operation));
  const text = preRows[index]?.text ?? "";
  const handle = handleOf(operation);
  if (handle) return [index, 1, handle.startOffset];
  switch (operation.type) {
    case "replaceInBlock":
      return [index, 1, Math.max(0, text.indexOf(String(operation["find"])))];
    case "splitBlock":
      return [index, 1, Number(operation["offset"])];
    case "mergeBlockWithNext":
    case "insertAfterBlock":
    case "insertTable":
      return [index + 1, 0, 0];
    case "setBlockParagraphProperties":
    case "insertBeforeBlock":
      return [index, 0, 0];
    default:
      return [index, 1, 0];
  }
};

const PARAGRAPH_PROPERTY_REPLAY = {
  replaceInBlock: "other",
  replaceRange: "other",
  commentOnRange: "other",
  formatRange: "other",
  insertAfterBlock: "other",
  insertBeforeBlock: "other",
  replaceBlock: "restyle",
  deleteBlock: "other",
  splitBlock: "other",
  mergeBlockWithNext: "other",
  setBlockParagraphProperties: "patch",
  insertTable: "other",
  deleteTable: "other",
  commentOnBlock: "other",
  insertSignatureTable: "other",
  insertTableRow: "other",
  deleteTableRow: "other",
  insertTableColumn: "other",
  deleteTableColumn: "other",
  mergeTableCells: "other",
  splitTableCell: "other",
} as const satisfies Record<
  FolioDocumentOperation["type"],
  "patch" | "restyle" | "other"
>;

const isOperationType = (
  type: string,
): type is FolioDocumentOperation["type"] =>
  Object.hasOwn(PARAGRAPH_PROPERTY_REPLAY, type);

type SequentialGroupsOptions = {
  applied: readonly AnyOperation[];
  preRows: readonly Row[];
  mode: Mode;
};

export const sequentialGroups = ({
  applied,
  preRows,
  mode,
}: SequentialGroupsOptions): AnyOperation[][] => {
  const indexed = applied.map((operation, index) => ({ operation, index }));
  const deleted = new Set(
    applied
      .filter((operation) => operation.type === "deleteBlock")
      .map(blockIdOf),
  );
  const containers = new Map<string, Row[]>();
  for (const row of preRows) {
    const key = row.table
      ? `${row.table.tableIndex}:${row.table.rowIndex}:${row.table.cellIndex}`
      : "story";
    const rows = containers.get(key) ?? [];
    rows.push(row);
    containers.set(key, rows);
  }
  const retiringCarriers = new Set<string>();
  for (const [container, rows] of containers) {
    const last = rows.at(-1);
    if (!last || (container === "story" && last !== preRows.at(-1))) continue;
    const appended = applied.some(
      (operation) =>
        operation.type === "insertAfterBlock" &&
        blockIdOf(operation) === last.id,
    );
    if (appended || !deleted.has(last.id)) continue;
    const carrier = rows.findLast((row) => !deleted.has(row.id));
    if (carrier) retiringCarriers.add(carrier.id);
  }
  const beforeFinalBreakRetirement = (operation: AnyOperation): boolean => {
    if (!isOperationType(operation.type)) {
      throw new Error(`Sequential replay has no policy for ${operation.type}`);
    }
    if (
      mode !== "tracked-changes" ||
      !retiringCarriers.has(blockIdOf(operation) ?? "")
    ) {
      return false;
    }
    const disposition = PARAGRAPH_PROPERTY_REPLAY[operation.type];
    switch (disposition) {
      case "patch":
        return true;
      case "restyle":
        return (
          operation["styleId"] !== undefined &&
          operation["styleId"] !==
            preRows.find((row) => row.id === blockIdOf(operation))?.styleId
        );
      case "other":
        return false;
      default: {
        const unhandled: never = disposition;
        throw new Error(
          `Unhandled paragraph property replay policy: ${unhandled}`,
        );
      }
    }
  };
  const earlyProperties = applied.filter(beforeFinalBreakRetirement);
  const earlyIds = new Set(earlyProperties.map(({ id }) => id));
  const isColumnEdit = (operation: AnyOperation): boolean =>
    operation.type === "insertTableColumn" ||
    operation.type === "deleteTableColumn";
  const deletedSplitTargets = new Set(
    applied
      .filter(
        (operation) =>
          mode === "tracked-changes" &&
          operation.type === "deleteBlock" &&
          applied.some(
            (candidate) =>
              candidate.type === "insertTable" &&
              candidate["position"] !== "before" &&
              blockIdOf(candidate) === blockIdOf(operation),
          ),
      )
      .map(blockIdOf),
  );
  const columnLocation = (operation: AnyOperation): [number, number] => {
    const targetId = blockIdOf(operation);
    const target = preRows.find((row) => row.id === targetId);
    if (!target?.table)
      throw new Error(
        `Applied ${operation.type} anchor ${targetId} has no source table.`,
      );
    const sourceColumn = target.table.gridColumnIndex;
    const column =
      operation.type === "insertTableColumn" &&
      operation["position"] !== "before"
        ? sourceColumn + target.table.columnSpan
        : sourceColumn;
    return [target.table.tableIndex, column];
  };
  const ordered = orderBatchOperations(
    indexed.filter(
      ({ operation }) =>
        !ANNOTATIONS.has(operation.type) && !earlyIds.has(operation.id),
    ),
    ({ operation, index }): BatchOrderAddress => {
      if (isColumnEdit(operation)) {
        const [tablePosition, columnIndex] = columnLocation(operation);
        return {
          type: "column",
          tablePosition,
          columnIndex,
          insertion: operation.type === "insertTableColumn",
          inputIndex: index,
        };
      }
      const [block, inside, offset] = placement(operation, preRows);
      const hasDeletedSplit = deletedSplitTargets.has(blockIdOf(operation));
      const position =
        hasDeletedSplit && operation.type === "deleteBlock"
          ? (block + 1) * 1_000_000_000
          : block * 1_000_000_000 + inside * 100_000_000 + offset;
      return {
        type: "content",
        position,
        rank: contentRank(operation, { hasDeletedSplit }),
        inputIndex: index,
      };
    },
  );
  const rest = ordered.filter(({ operation }) => !isColumnEdit(operation));
  const columns = ordered.filter(({ operation }) => isColumnEdit(operation));
  const inputOrder = new Map(applied.map(({ id }, index) => [id, index]));
  const groups = groupInsertionRuns(
    rest.map(({ operation }) => operation),
    (operation) => {
      if (!INSERTIONS.has(operation.type)) return null;
      const [block, inside, offset] = placement(operation, preRows);
      return block * 1_000_000_000 + inside * 100_000_000 + offset;
    },
  ).map((group) =>
    group.toSorted(
      (left, right) =>
        (inputOrder.get(left.id) ?? Number.MAX_SAFE_INTEGER) -
        (inputOrder.get(right.id) ?? Number.MAX_SAFE_INTEGER),
    ),
  );
  // A deletion removes or retires its block, so an insertion before it can
  // only be stated one at a time while the block still stands. The batch
  // keeps that gap, so the oracle runs the whole run ahead of the deletion.
  for (const { operation } of rest) {
    if (
      operation.type !== "insertBeforeBlock" ||
      !deleted.has(blockIdOf(operation))
    ) {
      continue;
    }
    const from = groups.findIndex((group) => group.includes(operation));
    const target = groups.findIndex(
      (group) =>
        group.length === 1 &&
        group[0]?.type === "deleteBlock" &&
        blockIdOf(group[0]) === blockIdOf(operation),
    );
    const moved = groups[from];
    if (moved && target >= 0 && target < from) {
      groups.splice(from, 1);
      groups.splice(target, 0, moved);
    }
  }
  const annotations = orderBatchOperations(
    indexed.filter(({ operation }) => ANNOTATIONS.has(operation.type)),
    ({ operation, index }): BatchOrderAddress => {
      const [position] = placement(operation, preRows);
      return { type: "annotation", position, inputIndex: index };
    },
  );
  const early = orderBatchOperations(
    indexed.filter(({ operation }) => earlyIds.has(operation.id)),
    ({ operation, index }): BatchOrderAddress => {
      const [block, inside, offset] = placement(operation, preRows);
      return {
        type: "content",
        position: block * 1_000_000_000 + inside * 100_000_000 + offset,
        rank: contentRank(operation, { hasDeletedSplit: false }),
        inputIndex: index,
      };
    },
  );
  return [
    ...annotations.map(({ operation }) => [operation]),
    ...early.map(({ operation }) => [operation]),
    ...groups,
    ...columns.map(({ operation }) => [operation]),
  ];
};
