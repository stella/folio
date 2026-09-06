/**
 * Runtime companions of the paint IR.
 *
 * `types.ts` is pure data by rule, so the handful of values a backend needs
 * live here instead: the discriminator list a coverage test enumerates, and
 * the two colours every producer resolves a theme variable to. Nothing here
 * carries a layout fact, which is why a backend may take a runtime edge to it
 * and not to `types.ts`.
 */

import type { DisplayColor, DisplayPrimitive, DisplayStrokePattern } from "./types";

type PrimitiveKind = DisplayPrimitive["kind"];

/**
 * Total map over the primitive kinds. `satisfies Record<PrimitiveKind, ...>`
 * is what makes it total in both directions: a new primitive kind that is not
 * listed fails to compile, and so does a name that is not a kind. A coverage
 * test built on this therefore cannot quietly stop covering a new kind.
 */
const PRIMITIVE_KIND_NAMES = {
  glyphRun: "glyphRun",
  rect: "rect",
  line: "line",
  image: "image",
  clipGroup: "clipGroup",
  rotateGroup: "rotateGroup",
  opacityGroup: "opacityGroup",
} as const satisfies Record<PrimitiveKind, PrimitiveKind>;

export const DISPLAY_PRIMITIVE_KINDS = Object.values(PRIMITIVE_KIND_NAMES);

/**
 * Dash geometry of each stroke pattern, as multiples of the stroke thickness.
 *
 * `null` means the pattern is not a dash: `solid` draws one unbroken line,
 * `double` two parallel lines, `wavy` a periodic curve. Sharing the factors is
 * what keeps a dashed underline in the editor and a dashed underline in the
 * export the same dash: with the pattern name alone, each backend would pick
 * its own period and nobody would notice until the two were laid side by side.
 */
export const STROKE_DASH_FACTORS = {
  solid: null,
  dashed: { dash: 3, gap: 2 },
  dotted: { dash: 1, gap: 1 },
  double: null,
  wavy: null,
} as const satisfies Record<DisplayStrokePattern, { dash: number; gap: number } | null>;

/**
 * Gap between the two lines of a `double` stroke, as a multiple of thickness.
 * Word draws the pair at the authored width with an equal gap between them.
 */
export const DOUBLE_STROKE_GAP_FACTOR = 1;

/**
 * Period and peak-to-peak amplitude of a `wavy` stroke, in thicknesses.
 *
 * The amplitude is measured between the centres of the curve at its extremes,
 * so the stroke's own thickness lies outside it: a wavy stroke's total cross
 * extent is `(WAVY_STROKE_AMPLITUDE_FACTOR + 1) * thicknessPx`, centred on the
 * path. Stated because "amplitude" alone leaves each backend to decide whether
 * the thickness is inside or outside, and the two answers differ by a
 * thickness on every wavy underline in the document.
 */
export const WAVY_STROKE_PERIOD_FACTOR = 6;
export const WAVY_STROKE_AMPLITUDE_FACTOR = 2;

export const BLACK: DisplayColor = { r: 0, g: 0, b: 0, a: 1 };
export const WHITE: DisplayColor = { r: 255, g: 255, b: 255, a: 1 };
