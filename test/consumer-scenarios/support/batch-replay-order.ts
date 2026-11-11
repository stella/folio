import {
  type BatchOrderAddress,
  contentRank,
  deletionRemovesBlock,
  groupInsertionRuns,
  orderAddressingBeforeRemoval,
  orderBatchOperations,
} from "@stll/docx-core/ops";

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
  /**
   * Whether the batch places a block in the gap before a deleted final
   * paragraph's run: the paragraph and the ones before it in its container the
   * batch deletes too. Whether the paragraph goes or stays emptied is read
   * from what precedes it once the batch is finished, in every mode, so its
   * deletion is replayed after such a block lands.
   */
  const placesBeforeRun = (last: Row): boolean => {
    const lastAt = preRows.indexOf(last);
    let start = lastAt;
    while (start > 0) {
      const row = preRows.at(start - 1);
      if (!sameContainerRows(last, row) || !deleted.has(row.id)) break;
      start -= 1;
    }
    return applied.some((operation) => {
      if (!INSERTIONS.has(operation.type)) return false;
      const anchor = preRows.find((row) => row.id === blockIdOf(operation));
      if (anchor === undefined) return false;
      const [index] = placement(operation, preRows);
      if (index < start || index > lastAt) return false;
      // A paragraph placed beside a cell goes outside the table, into the
      // container the table stands in.
      return anchor.table !== undefined && last.table === undefined
        ? operation.type !== "insertTable"
        : sameContainerRows(last, anchor);
    });
  };
  const retiringFinals = new Set<string>();
  const finalBlocks = new Set<string>();
  const uncarriedFinals = new Set<string>();
  // A suggestion reads the live rows, where a paragraph pending deletion is a
  // blank block: the block before such blocks ends its container as the batch
  // sees it.
  const finalBeforeBlanks = new Set<string>();
  for (const [container, rows] of containers) {
    const last = rows.at(-1);
    if (!last || (container === "story" && last !== preRows.at(-1))) continue;
    finalBlocks.add(last.id);
    const beforeBlanks = rows.findLast(({ text }) => text.length > 0);
    if (mode === "suggested" && beforeBlanks) {
      finalBeforeBlanks.add(beforeBlanks.id);
    }
    const appended = applied.some(
      (operation) =>
        operation.type === "insertAfterBlock" &&
        blockIdOf(operation) === last.id,
    );
    if (appended || !deleted.has(last.id)) continue;
    const carrier = breakCarrierOf(last);
    if (!carrier) uncarriedFinals.add(last.id);
    // A sole paragraph with one inserted before it is deleted first, and the
    // insertion takes its place. A paragraph that ends a deleted run stays until
    // the run is gone. Applied directly only such a replacement waits, since
    // nothing before it holds a break to retire.
    const replaced = applied.some(
      (operation) =>
        operation.type === "insertBeforeBlock" &&
        blockIdOf(operation) === last.id,
    );
    const lastAt = preRows.indexOf(last);
    const endsARun =
      lastAt > 0 && sameContainerRows(last, preRows.at(lastAt - 1));
    const runsLate =
      mode === "tracked-changes"
        ? !replaced || endsARun
        : mode === "direct" && carrier === undefined && replaced && endsARun;
    if (runsLate || placesBeforeRun(last)) retiringFinals.add(last.id);
  }
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
    indexed.filter(({ operation }) => !ANNOTATIONS.has(operation.type)),
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
  // A container's final paragraph differs. With no surviving paragraph to
  // retire its mark onto, a paragraph inserted first would carry the break,
  // or take the deleted paragraph's place directly, and the batch deletes
  // first. In a suggestion the batch hands the break to a paragraph inserted
  // into the gap before it, whichever side of the gap the insertion is stated
  // from.
  const gapDeletion = (operation: AnyOperation): string | undefined => {
    const id = blockIdOf(operation);
    if (operation.type === "insertBeforeBlock") {
      const retires = mode !== "suggested" && uncarriedFinals.has(id);
      return deleted.has(id) && !retires ? id : undefined;
    }
    if (operation.type !== "insertAfterBlock" || mode === "direct") {
      return undefined;
    }
    const index = preRows.findIndex((row) => row.id === id);
    const next = preRows.at(index + 1);
    return next !== undefined &&
      deleted.has(next.id) &&
      (finalBlocks.has(next.id) || finalBeforeBlanks.has(next.id)) &&
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
    // The deletion moves, not the run: what else is at the gap keeps its place
    // beside the run, in the order the batch runs them.
    const deletion = groups[target];
    if (deletion && target >= 0 && target < from) {
      groups.splice(target, 1);
      groups.splice(from, 0, deletion);
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
    // Removing a table row or table takes its anchor with it.
    if (operation.type === "deleteTableRow" || operation.type === "deleteTable")
      return [id];
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
  // The batch retires a final paragraph's break once, over the finished batch,
  // when what precedes it is settled. Until then the paragraphs before it can
  // still be written to and placed beside, so its deletion runs late. A merge
  // already reads that break as gone, so the deletion runs before any merge.
  const retiresBreak = (group: readonly AnyOperation[]): boolean =>
    group.every(
      (operation) =>
        operation.type === "deleteBlock" &&
        retiringFinals.has(blockIdOf(operation)),
    );
  const others = reordered.filter((group) => !retiresBreak(group));
  const firstMerge = others.findIndex((group) =>
    group.some(({ type }) => type === "mergeBlockWithNext"),
  );
  const retirementsAt = firstMerge < 0 ? others.length : firstMerge;
  return [
    ...annotations.map(({ operation }) => [operation]),
    ...others.slice(0, retirementsAt),
    ...reordered.filter(retiresBreak),
    ...others.slice(retirementsAt),
    ...columns.map(({ operation }) => [operation]),
  ];
};
