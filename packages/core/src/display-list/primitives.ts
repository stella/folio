/**
 * Runtime companions of the paint IR.
 *
 * `types.ts` is pure data by rule, so what a backend needs at runtime lives
 * here instead: the discriminator list a coverage test enumerates, the two
 * colours every producer resolves a theme variable to, the stroke geometry
 * both backends draw from, and the rule that turns a run's advances into a
 * position per code point. A rule two backends must agree on belongs in one
 * module, not in one copy each. Nothing here carries a layout fact, which is
 * why a backend may take a runtime edge to it and not to `types.ts`.
 */

import type {
  DisplayColor,
  DisplayGlyphRun,
  DisplayPrimitive,
  DisplayStrokePattern,
} from "./types";

/**
 * Cursive joining, for the backends.
 *
 * A paint module needs a shaping predicate because a backend that places code
 * points independently has to know which ones may not be separated: in Arabic,
 * Syriac, N'Ko and Adlam a letter's glyph depends on its neighbours, so a
 * letter cut into a box of its own paints in isolated form and the word comes
 * apart where a reader sees it. Joining types are Unicode data rather than a
 * layout fact, which is what lets a backend take a runtime edge to them here.
 */
export { hasCursiveLetter, joinsAcrossBoundary } from "../utils/cursiveJoining";

type PrimitiveKind = DisplayPrimitive["kind"];

/**
 * Left edge of every code point of a run, as an offset from the run's `xPx`.
 *
 * One implementation for both backends. The ordering rule is prose in
 * `types.ts`, and two backends reading that prose separately is the
 * divergence the display list exists to prevent: a run occupies
 * `[xPx, xPx + sum(advancesPx)]` whichever way it runs, so under `ltr` code
 * point `i` starts at `sum(advancesPx[0..i))` and under `rtl` it ends there,
 * counted back from the run's right edge. A backend adds its own origin: the
 * left edge is where a glyph is placed in either direction, because a glyph
 * advances rightward from its origin whatever the paragraph does.
 */
export const glyphCellOffsetsPx = (run: DisplayGlyphRun): readonly number[] => {
  const totalPx = run.advancesPx.reduce((total, advance) => total + advance, 0);
  const offsetsPx: number[] = [];
  let consumedPx = 0;
  for (const advancePx of run.advancesPx) {
    offsetsPx.push(run.direction === "ltr" ? consumedPx : totalPx - consumedPx - advancePx);
    consumedPx += advancePx;
  }
  return offsetsPx;
};

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
