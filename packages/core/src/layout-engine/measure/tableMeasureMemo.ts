/**
 * The measure of a table that a layout pass handed out again, or moved.
 *
 * A table an edit did not reach is handed out by the next pass as the same
 * block, or, when the edit was before it, as a copy with shifted positions
 * (`shiftFlowBlockPmAnchors`). Neither changes what measuring it reads: a
 * block handed out is never written again, and measures hold no positions
 * (`MEASURES_HOLD_NO_POSITIONS`). So the measure taken for the block, or for
 * the block it was moved from, holds at the same width, while the caches
 * measuring reads from stay as they were.
 */

import type { TableBlock, TableMeasure } from "../types";
import { getLineBreakProviderGeneration } from "./lineBreakProvider";

type Taken = {
  measure: TableMeasure;
  width: number;
  generation: number;
  lineBreaks: number;
};

/** The block each moved table was moved from. */
const origins = new WeakMap<TableBlock, TableBlock>();
const taken = new WeakMap<TableBlock, Taken>();
/** Advanced whenever the measuring caches are cleared. */
let generation = 0;

/**
 * Record that `copy` measures as `origin` does: it is `origin` moved, or
 * `origin` with values derived from `origin` alone. Copies of copies lead to
 * the first block, so a measure recorded for any of them serves them all.
 */
export const noteTableOrigin = (copy: TableBlock, origin: TableBlock): void => {
  origins.set(copy, origins.get(origin) ?? origin);
};

/** The measure taken for `block` or the block it was moved from, at `width`. */
export const takenTableMeasure = (
  block: TableBlock,
  width: number,
): TableMeasure | undefined => {
  const own = taken.get(block);
  const origin = origins.get(block);
  const entry = own ?? (origin === undefined ? undefined : taken.get(origin));
  if (
    entry === undefined ||
    entry.width !== width ||
    entry.generation !== generation ||
    entry.lineBreaks !== getLineBreakProviderGeneration()
  )
    return undefined;
  if (own === undefined) taken.set(block, entry);
  return entry.measure;
};

export const recordTableMeasure = (
  block: TableBlock,
  width: number,
  measure: TableMeasure,
): void => {
  const entry = {
    measure,
    width,
    generation,
    lineBreaks: getLineBreakProviderGeneration(),
  };
  taken.set(block, entry);
  const origin = origins.get(block);
  if (origin !== undefined) taken.set(origin, entry);
};

/** Forget every table measure: what measuring reads from has changed. */
export const forgetTableMeasures = (): void => {
  generation += 1;
};
