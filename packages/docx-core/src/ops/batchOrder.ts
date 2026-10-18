/**
 * The execution coordinates of an operation resolved against a batch's
 * pre-state. Positions and table coordinates are ordered from right to left;
 * equal coordinates use reverse input order. Annotations run before moving
 * edits, while table-shape and column edits run after ordinary content edits.
 */
export type BatchOrderAddress =
  | { type: "annotation"; position: number; inputIndex: number }
  | { type: "content"; position: number; inputIndex: number }
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

/** Compare two pre-batch addresses in the documented batch execution order. */
export const compareBatchOrder = (left: BatchOrderAddress, right: BatchOrderAddress): number => {
  if (left.type !== right.type) return PHASE[left.type] - PHASE[right.type];
  if (
    (left.type === "annotation" || left.type === "content") &&
    (right.type === "annotation" || right.type === "content")
  ) {
    return right.position - left.position || right.inputIndex - left.inputIndex;
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
    if (left.insertion !== right.insertion) insertionOrder = left.insertion ? -1 : 1;
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
  operations.toSorted((left, right) => compareBatchOrder(addressOf(left), addressOf(right)));
