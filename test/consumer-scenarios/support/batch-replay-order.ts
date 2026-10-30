import {
  type BatchOrderAddress,
  contentRank,
  deletionRemovesBlock,
  groupInsertionRuns,
  orderAddressingBeforeRemoval,
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
  // A paragraph added beside a table cell goes outside the table: before its
  // first row or after its last, so every anchor in one outermost table shares a gap.
  const outer = preRows[index]?.table?.outerTableIndex;
  if (outer !== undefined) {
    const rows = preRows.flatMap((row, at) =>
      row.table?.outerTableIndex === outer ? [at] : [],
    );
    if (operation.type === "insertBeforeBlock") {
      return [rows.at(0) ?? index, 0, 0];
    }
    if (operation.type === "insertAfterBlock") {
      return [(rows.at(-1) ?? index) + 1, 0, 0];
    }
  }
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

const sameContainerRows = (left: Row, right: Row | undefined): right is Row =>
  right !== undefined &&
  `${left.table?.tableIndex}:${left.table?.rowIndex}:${left.table?.cellIndex}` ===
    `${right.table?.tableIndex}:${right.table?.rowIndex}:${right.table?.cellIndex}`;

type SequentialGroupsOptions = {
  applied: readonly AnyOperation[];
  /** The rows the batch's outcome is read against (accepted, for tracked changes). */
  preRows: readonly Row[];
  /**
   * The live rows the operations name, when `preRows` are not those. A
   * column edit may anchor in a cell of a row pending deletion, which only
   * the live table still has; column edits order among themselves, so they
   * all take live coordinates.
   */
  liveRows?: readonly Row[];
  mode: Mode;
};

export const sequentialGroups = ({
  applied,
  preRows,
  liveRows = preRows,
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
  /**
   * The paragraph whose break a deleted container-final paragraph retires:
   * the nearest earlier one of its container the batch keeps, with only
   * paragraphs the batch deletes between them. A table between them (rows of
   * another container, `preRows` being in document order) leaves none, and
   * the final paragraph stays emptied, as the engine's
   * `finalParagraphPredecessor` has it.
   */
  const breakCarrierOf = (last: Row): Row | undefined => {
    for (let at = preRows.indexOf(last) - 1; at >= 0; at -= 1) {
      const row = preRows.at(at);
      if (!sameContainerRows(last, row)) return undefined;
      if (!deleted.has(row.id)) return row;
    }
    return undefined;
  };
  const retiringCarriers = new Set<string>();
  const finalBlocks = new Set<string>();
  const uncarriedFinals = new Set<string>();
  for (const [container, rows] of containers) {
    const last = rows.at(-1);
    if (!last || (container === "story" && last !== preRows.at(-1))) continue;
    finalBlocks.add(last.id);
    const appended = applied.some(
      (operation) =>
        operation.type === "insertAfterBlock" &&
        blockIdOf(operation) === last.id,
    );
    if (appended || !deleted.has(last.id)) continue;
    const carrier = breakCarrierOf(last);
    if (carrier) retiringCarriers.add(carrier.id);
    else uncarriedFinals.add(last.id);
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
    const target = liveRows.find((row) => row.id === targetId);
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
    (operation) => operation.type === "setBlockParagraphProperties",
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
  // A container's final paragraph differs. In tracked changes, with no
  // surviving paragraph to retire its mark onto, a paragraph inserted first
  // would carry the break and the batch deletes first. In a suggestion the batch hands
  // the break to a paragraph inserted into the gap before it, whichever side
  // of the gap the insertion is stated from.
  const gapDeletion = (operation: AnyOperation): string | undefined => {
    const id = blockIdOf(operation);
    if (operation.type === "insertBeforeBlock") {
      const retires = mode === "tracked-changes" && uncarriedFinals.has(id);
      return deleted.has(id) && !retires ? id : undefined;
    }
    if (operation.type !== "insertAfterBlock" || mode === "direct") {
      return undefined;
    }
    const index = preRows.findIndex((row) => row.id === id);
    const next = preRows.at(index + 1);
    return next !== undefined &&
      deleted.has(next.id) &&
      finalBlocks.has(next.id) &&
      sameContainerRows(next, preRows.at(index))
      ? next.id
      : undefined;
  };
  for (const { operation } of rest) {
    const deletedId = gapDeletion(operation);
    if (deletedId === undefined) continue;
    const from = groups.findIndex((group) => group.includes(operation));
    const target = groups.findIndex(
      (group) =>
        group.length === 1 &&
        group[0]?.type === "deleteBlock" &&
        blockIdOf(group[0]) === deletedId,
    );
    const moved = groups[from];
    if (moved && target >= 0 && target < from) {
      groups.splice(from, 1);
      groups.splice(target, 0, moved);
    }
  }
  // The shared batch order: a run addressing a block goes before the run that
  // deletes or merges it away, as the engine executes the batch.
  const appendedAfter = (id: string) =>
    applied.some(
      (operation) =>
        blockIdOf(operation) === id &&
        (operation.type === "insertAfterBlock" ||
          (operation.type === "insertTable" &&
            operation["position"] !== "before")),
    );
  const removedBy = (operation: AnyOperation): string[] => {
    const id = blockIdOf(operation);
    if (operation.type === "deleteBlock") {
      const row = preRows.find((candidate) => candidate.id === id);
      return row === undefined ||
        deletionRemovesBlock({
          endsContainer: finalBlocks.has(id) && !appendedAfter(id),
          hasKeptParagraphBefore: breakCarrierOf(row) !== undefined,
        })
        ? [id]
        : [];
    }
    if (operation.type !== "mergeBlockWithNext") return [];
    const index = preRows.findIndex((row) => row.id === id);
    const row = preRows.at(index);
    const next = preRows.at(index + 1);
    return row !== undefined && sameContainerRows(row, next) ? [next.id] : [];
  };
  const reordered = orderAddressingBeforeRemoval(groups, (group) => ({
    addresses: group.flatMap((operation) =>
      operation.type === "deleteBlock" ||
      operation.type === "deleteTableRow" ||
      operation.type === "deleteTable"
        ? []
        : [blockIdOf(operation)],
    ),
    removes: group.flatMap(removedBy),
  }));
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
    ...reordered,
    ...columns.map(({ operation }) => [operation]),
  ];
};
