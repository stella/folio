import type { Transaction } from "prosemirror-state";
import { AddMarkStep, RemoveMarkStep } from "prosemirror-transform";

import { mergeDirtyRanges, type DirtyRange } from "./incrementalMeasure";

type DirtyRangeAccumulator = {
  from: number;
  to: number;
};

export function getTransactionDirtyRange(transaction: Transaction): DirtyRange | null {
  const dirtyRange = {
    from: Number.POSITIVE_INFINITY,
    to: Number.NEGATIVE_INFINITY,
  };

  for (let stepIndex = 0; stepIndex < transaction.mapping.maps.length; stepIndex += 1) {
    includeStepDirtyRange(transaction, stepIndex, dirtyRange);
  }

  if (!Number.isFinite(dirtyRange.from) || !Number.isFinite(dirtyRange.to)) {
    return null;
  }

  return dirtyRange;
}

/** Collect every transaction in an applied batch in the final document's coordinates. */
export function getTransactionsDirtyRange(transactions: readonly Transaction[]): DirtyRange | null {
  let combinedRange: DirtyRange | null = null;

  for (let index = 0; index < transactions.length; index += 1) {
    const transaction = transactions[index];
    if (!transaction) {
      continue;
    }
    const range = getTransactionDirtyRange(transaction);
    if (!range) {
      continue;
    }

    let from = range.from;
    let to = range.to;
    for (const followingTransaction of transactions.slice(index + 1)) {
      // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map(pos, assoc) API
      from = followingTransaction.mapping.map(from, -1);
      // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map(pos, assoc) API
      to = followingTransaction.mapping.map(to, 1);
    }
    combinedRange = mergeDirtyRanges(combinedRange, { from, to });
  }

  return combinedRange;
}

function includeStepDirtyRange(
  transaction: Transaction,
  stepIndex: number,
  dirtyRange: DirtyRangeAccumulator,
): void {
  const map = transaction.mapping.maps[stepIndex];
  if (!map) {
    return;
  }

  const followingMaps = transaction.mapping.slice(stepIndex + 1);
  const extend = (newStart: number, newEnd: number) => {
    // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map(pos, assoc) API
    const finalStart = followingMaps.map(newStart, -1);
    // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map(pos, assoc) API
    const finalEnd = followingMaps.map(newEnd, 1);
    dirtyRange.from = Math.min(dirtyRange.from, finalStart, finalEnd);
    dirtyRange.to = Math.max(dirtyRange.to, finalStart, finalEnd);
  };

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap API
  map.forEach((_oldStart, _oldEnd, newStart, newEnd) => extend(newStart, newEnd));

  // AddMarkStep / RemoveMarkStep produce an empty StepMap (mark changes don't
  // move positions), so the loop above misses them and the incremental
  // measure path falls back to a full re-measure. Read the affected range
  // directly off the step so mark-only edits stay incremental.
  const step = transaction.steps[stepIndex];
  if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
    extend(step.from, step.to);
  }
}
