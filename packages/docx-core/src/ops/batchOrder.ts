/**
 * The execution coordinates of an operation resolved against a batch's
 * pre-state. Positions and table coordinates are ordered from right to left;
 * equal content coordinates use an explicit rank, then reverse input order.
 * Annotations run before moving edits, while table-shape and column edits run
 * after ordinary content edits. Table mutations in nested tables still run
 * right to left before the phase order is applied within the same table.
 */
export type BatchOrderAddress =
  | { type: "annotation"; position: number; inputIndex: number }
  | { type: "content"; position: number; rank: number; inputIndex: number }
  | {
      type: "cell";
      tablePosition: number;
      bottom: number;
      right: number;
      inputIndex: number;
    }
  | {
      type: "column";
      tablePosition: number;
      columnIndex: number;
      insertion: boolean;
      inputIndex: number;
    };

const PHASE = { annotation: 0, content: 1, cell: 2, column: 3 } as const;

type ContentRankOperation = { type: string; position?: unknown };

/**
 * Rank content operations at one pre-batch target. A tracked split deletion
 * must retract the inserted paragraph before a table is placed after it;
 * ordinary edits stay between those two exceptional operations.
 */
export const contentRank = (
  operation: ContentRankOperation,
  target: { hasDeletedSplit: boolean },
): number => {
  if (!target.hasDeletedSplit) return 1;
  if (operation.type === "deleteBlock") return 0;
  if (operation.type === "insertTable" && operation.position !== "before")
    return 2;
  return 1;
};

/** Compare two pre-batch addresses in the documented batch execution order. */
export const compareBatchOrder = (
  left: BatchOrderAddress,
  right: BatchOrderAddress,
): number => {
  if (
    (left.type === "cell" || left.type === "column") &&
    (right.type === "cell" || right.type === "column") &&
    left.tablePosition !== right.tablePosition
  ) {
    return right.tablePosition - left.tablePosition;
  }
  if (left.type !== right.type) return PHASE[left.type] - PHASE[right.type];
  if (
    (left.type === "annotation" || left.type === "content") &&
    (right.type === "annotation" || right.type === "content")
  ) {
    return (
      right.position - left.position ||
      (left.type === "content" && right.type === "content"
        ? left.rank - right.rank
        : 0) ||
      right.inputIndex - left.inputIndex
    );
  }
  if (left.type === "cell" && right.type === "cell") {
    return (
      right.tablePosition - left.tablePosition ||
      right.bottom - left.bottom ||
      right.right - left.right ||
      right.inputIndex - left.inputIndex
    );
  }
  if (left.type === "column" && right.type === "column") {
    let insertionOrder = 0;
    if (left.insertion !== right.insertion)
      insertionOrder = left.insertion ? -1 : 1;
    return (
      right.tablePosition - left.tablePosition ||
      right.columnIndex - left.columnIndex ||
      insertionOrder ||
      right.inputIndex - left.inputIndex
    );
  }
  return 0;
};

/** Order operations once from addresses captured before the batch changes the document. */
export const orderBatchOperations = <T>(
  operations: readonly T[],
  addressOf: (operation: T) => BatchOrderAddress,
): T[] =>
  operations.toSorted((left, right) =>
    compareBatchOrder(addressOf(left), addressOf(right)),
  );

/**
 * Partition an ordered batch into execution runs. Only consecutive insertions
 * at the same position share a run; every other operation remains a singleton.
 */
export const groupInsertionRuns = <T>(
  operations: readonly T[],
  insertionPositionOf: (operation: T) => number | null,
): T[][] => {
  const groups: T[][] = [];
  for (const operation of operations) {
    const position = insertionPositionOf(operation);
    const previous = groups.at(-1);
    const first = previous?.at(0);
    if (
      position !== null &&
      previous !== undefined &&
      first !== undefined &&
      insertionPositionOf(first) === position
    ) {
      previous.push(operation);
      continue;
    }
    groups.push([operation]);
  }
  return groups;
};
