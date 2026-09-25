/**
 * Shared by the document, paragraph, and table layout passes: the rendered-break
 * reflow tolerance and footnote reserve growth on the current page.
 */

import type { PageState } from "./paginator";

export const RENDERED_BREAK_REFLOW_TOLERANCE_LINES = 3;

export function projectedFootnoteReserveGrowth(
  state: PageState,
  additionalDemandHeight: number,
): number {
  const projectedDemand = state.footnoteDemandHeight + additionalDemandHeight;
  return Math.max(state.footnoteHeightFloor, projectedDemand) - state.footnoteHeight;
}
