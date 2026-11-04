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

/** The pre-batch blocks an operation names and the ones it takes away. */
export type BatchBlockAccess<Block> = {
  /** Blocks the operation edits or places something beside. */
  addresses: readonly Block[];
  /** Blocks the operation deletes, or merges into the block before them. */
  removes: readonly Block[];
};

/**
 * Whether deleting a block takes it out of the document. A container (a
 * story, cell, note or text box) has to end with a paragraph, so deleting its
 * last one with no paragraph the batch keeps before it to take its break only
 * empties it: it stays, and can still be addressed.
 */
export const deletionRemovesBlock = (deletion: {
  endsContainer: boolean;
  hasKeptParagraphBefore: boolean;
}): boolean => !deletion.endsContainer || deletion.hasKeptParagraphBefore;

/**
 * A batch equals its operations applied one at a time with every operation
 * that addresses a block before any operation that removes that block: a
 * single edit can only address a block that still stands. Each operation
 * that addresses a block another one ahead of it removes moves to just before
 * the earliest such remover; every other operation keeps its place.
 */
export const orderAddressingBeforeRemoval = <T, Block>(
  ordered: readonly T[],
  accessOf: (operation: T) => BatchBlockAccess<Block>,
): T[] => {
  const result = [...ordered];
  const removesAny = (remover: T, blocks: readonly Block[]) => {
    const { removes } = accessOf(remover);
    return blocks.some((block) => removes.includes(block));
  };
  // Every move takes an operation earlier, past at least one remover, so the
  // order settles; the bound only guards against a cycle.
  for (let round = 0; round <= result.length * result.length; round += 1) {
    const from = result.findIndex((operation, at) => {
      const { addresses } = accessOf(operation);
      return (
        addresses.length > 0 &&
        result.slice(0, at).some((earlier) => removesAny(earlier, addresses))
      );
    });
    if (from < 0) return result;
    const operation = result[from];
    if (operation === undefined) return result;
    const { addresses } = accessOf(operation);
    const to = result.findIndex((candidate) =>
      removesAny(candidate, addresses),
    );
    result.splice(from, 1);
    result.splice(to, 0, operation);
  }
  throw new Error(
    "A batch has no order that addresses each block before removing it",
  );
};

/**
 * Partition an ordered batch into execution runs. Insertions at the same
 * position share a run; every other operation remains a singleton.
 *
 * An operation that `keepsGap` admits (one that writes properties and places
 * nothing in the gap) does not end a run: an equal-gap insertion after it
 * still joins the run, which runs before it. Otherwise a property write tied
 * between two insertions at one gap would split them, and they would inherit
 * their anchor's properties from different moments while landing in one
 * requested sequence, which no series of single edits produces.
 */
export const groupInsertionRuns = <T>(
  operations: readonly T[],
  insertionPositionOf: (operation: T) => number | null,
  keepsGap: (operation: T) => boolean = () => false,
): T[][] => {
  const groups: T[][] = [];
  let open: { run: T[]; position: number } | null = null;
  for (const operation of operations) {
    const position = insertionPositionOf(operation);
    if (position !== null && open !== null && open.position === position) {
      open.run.push(operation);
      continue;
    }
    if (position !== null) {
      const run = [operation];
      groups.push(run);
      open = { run, position };
      continue;
    }
    groups.push([operation]);
    if (!keepsGap(operation)) open = null;
  }
  return groups;
};
