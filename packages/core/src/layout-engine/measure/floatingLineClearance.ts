/**
 * Finds where a line clears floating objects that leave too little width
 * beside them.
 */

import {
  getFloatingAvailableWidth,
  getFloatingMargins,
  type FloatingImageZone,
} from "./floatingZones";

/**
 * Minimum horizontal room a line must offer before we treat it as usable for
 * body text. Below this threshold the line is bumped past obstructing floats
 * via `findClearLineY` instead of being rendered into the unusable sliver.
 * Without this guard, a near-full-width float (e.g. floating table) produces
 * a ~2px segment that collapses every wrap line to one glyph per row.
 */
export const MIN_WRAP_SEGMENT_WIDTH = 24;

/**
 * Find the next vertical position at or below `startY` where the available
 * text width is at least `minWidth`. Used to skip lines past stacked floats
 * when there is no horizontal room for meaningful text at the current Y.
 *
 * Returns `startY` if the current position already has enough room, otherwise
 * the lowest `bottomY` of any zone currently obstructing the line. The caller
 * is expected to re-query margins at the returned Y.
 *
 * Coordinates are absolute (i.e., already include any paragraphYOffset).
 */
export function findClearLineY(
  startY: number,
  lineHeight: number,
  zones: FloatingImageZone[] | undefined,
  contentWidth: number,
  minWidth: number,
): number {
  if (!zones || zones.length === 0) {
    return startY;
  }

  let y = startY;
  // Bounded loop — at most one step per zone the line currently overlaps,
  // plus a safety cushion. Prevents pathological re-entry while keeping the
  // happy path O(zones).
  for (let i = 0; i < zones.length + 2; i++) {
    const margins = getFloatingMargins(y, lineHeight, zones, 0);
    const available = Math.max(0, getFloatingAvailableWidth(margins, contentWidth));
    if (available >= minWidth) {
      return y;
    }

    const lineBottom = y + lineHeight;
    let nextY = Number.POSITIVE_INFINITY;
    for (const zone of zones) {
      // Skip zones we are already past or that lie entirely below this line.
      if (lineBottom <= zone.topY || y >= zone.bottomY) {
        continue;
      }
      if (zone.bottomY > y && zone.bottomY < nextY) {
        nextY = zone.bottomY;
      }
    }
    if (!Number.isFinite(nextY) || nextY <= y) {
      return y;
    }
    y = nextY;
  }
  return y;
}
