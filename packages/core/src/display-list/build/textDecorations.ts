/**
 * Text-decoration geometry.
 *
 * The DOM painter delegates underlines and strikethroughs to CSS
 * `text-decoration`, which derives its position and thickness from the font's
 * own `post`/`OS/2` tables. A display list has to state the geometry, and the
 * builder has no font binary, so these ratios stand in for the font-derived
 * metrics.
 *
 * They are approximations of what a browser computes from the face, chosen to
 * match the common Latin text faces at body sizes. The follow-up is to read
 * `underlinePosition` / `underlineThickness` from the embedded face and thread
 * them through the measure seam alongside ascent and descent; until then a
 * decoration on an unusual face sits a fraction of an em off where Word puts
 * it. Every constant is in one place so that follow-up is a single edit.
 */

/** Distance below the baseline to the centre of the underline, as a fraction of the em. */
export const UNDERLINE_OFFSET_RATIO = 0.1;

/** Underline thickness as a fraction of the em, floored so a hairline stays visible. */
export const UNDERLINE_THICKNESS_RATIO = 1 / 14;
export const MIN_DECORATION_THICKNESS_PX = 1;

/** Distance above the baseline to the centre of the strikethrough, as a fraction of the em. */
export const STRIKETHROUGH_OFFSET_RATIO = 0.28;

export const underlineThicknessPx = (fontSizePx: number): number =>
  Math.max(MIN_DECORATION_THICKNESS_PX, fontSizePx * UNDERLINE_THICKNESS_RATIO);

export const underlineCenterYPx = (baselineYPx: number, fontSizePx: number): number =>
  baselineYPx + fontSizePx * UNDERLINE_OFFSET_RATIO;

export const strikethroughCenterYPx = (baselineYPx: number, fontSizePx: number): number =>
  baselineYPx - fontSizePx * STRIKETHROUGH_OFFSET_RATIO;
