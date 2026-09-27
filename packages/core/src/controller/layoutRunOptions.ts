import type { LayoutRunReason } from "../layout-engine/layoutInstrumentation";

/**
 * How a layout pass measures its blocks.
 *
 * - `full` measures every block.
 * - `incremental` may keep the committed layout's measures for the blocks the
 *   current document shares with the document that layout was computed for.
 *   The pass derives what changed from those two documents itself; a caller
 *   never supplies a changed range, so the range cannot be relative to a
 *   document the committed measures did not come from.
 */
export const LAYOUT_MEASURE = {
  full: "full",
  incremental: "incremental",
} as const;

export type LayoutMeasure = (typeof LAYOUT_MEASURE)[keyof typeof LAYOUT_MEASURE];

export type LayoutRunOptions = {
  /** Defaults to {@link LAYOUT_MEASURE.full}. */
  measure?: LayoutMeasure;
  reason?: LayoutRunReason;
};
