import type { Transaction } from "prosemirror-state";
import { AddMarkStep, RemoveMarkStep } from "prosemirror-transform";

import { sweepPositions, type PositionQuery } from "../prosemirror/positionSweep";
import type { DirtyRange } from "./incrementalMeasure";

export function getTransactionDirtyRange(transaction: Transaction): DirtyRange | null {
  // Each step's changed range is carried through the steps after it. The sweep
  // maps them all in one pass; slicing the mapping per step would walk the
  // rest of the transaction once per step.
  const queries: PositionQuery[] = [];
  const extend = (newStart: number, newEnd: number, from: number): void => {
    queries.push({ pos: newStart, assoc: -1, from }, { pos: newEnd, assoc: 1, from });
  };

  for (let stepIndex = 0; stepIndex < transaction.mapping.maps.length; stepIndex += 1) {
    const map = transaction.mapping.maps[stepIndex];
    if (!map) {
      continue;
    }
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap API
    map.forEach((_oldStart, _oldEnd, newStart, newEnd) => extend(newStart, newEnd, stepIndex + 1));

    // AddMarkStep / RemoveMarkStep produce an empty StepMap (mark changes don't
    // move positions), so the loop above misses them and the incremental
    // measure path falls back to a full re-measure. Read the affected range
    // directly off the step so mark-only edits stay incremental.
    const step = transaction.steps[stepIndex];
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      extend(step.from, step.to, stepIndex + 1);
    }
  }

  if (queries.length === 0) {
    return null;
  }

  let from = Number.POSITIVE_INFINITY;
  let to = Number.NEGATIVE_INFINITY;
  for (const { pos } of sweepPositions([transaction.mapping], queries)) {
    from = Math.min(from, pos);
    to = Math.max(to, pos);
  }
  return { from, to };
}
