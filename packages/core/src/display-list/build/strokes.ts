/**
 * Border, outline and decoration strokes.
 *
 * One deliberate divergence from `layout-painter/borderStroke.ts`: that helper
 * encodes a sub-pixel border width as a `color-mix` alpha because CSS quantizes
 * a fractional border to one device pixel. A display list has no device pixel,
 * and a PDF strokes 0.4pt natively, so the authored fractional width and the
 * plain colour survive here. Everything else (colour/width defaults) is taken
 * from the painter's helper so the two cannot drift.
 *
 * Three vocabularies reach this module and each one has its own table. They
 * shared a single string-keyed lookup before, which is why `dash` resolved the
 * same whether it arrived as a CSS `border-style`, a DrawingML preset dash or
 * a `text-decoration-style`, and why every member no vocabulary spelled the
 * same way (`lgDashDot`, `sysDot`, `dottedHeavy`, …) fell through to a plain
 * line without anything able to say which table was short. Each table is total
 * over its own union at compile time, so a member can only be added with a
 * decision attached.
 */

import type { PresetLineDashVal, UnderlineStyle } from "@stll/docx-core/model";

import { resolveCssBorderStroke } from "../../layout-painter/borderStroke";
import type { CssBorderStyle } from "../../utils/borderCss";
import type { DisplayStroke, DisplayStrokePattern } from "../types";
import { parseDisplayColor } from "./colors";
import { HEAVY_UNDERLINE_WEIGHT, PLAIN_UNDERLINE_WEIGHT } from "./textDecorations";

/** A pattern, or the absence of a line: an edge that paints nothing. */
type StrokePattern = DisplayStrokePattern | "none";

type BorderInput = {
  color?: string | undefined;
  style?: CssBorderStyle | undefined;
  width?: number | undefined;
};

/**
 * Every CSS `border-style` folio paints, as a display-list pattern.
 *
 * The bevelled keywords have no display-list pattern, and Word draws
 * `threeDEmboss`/`threeDEngrave` as a plain line where the shading is
 * unavailable, so they stroke solid.
 */
export const CSS_BORDER_STROKE_PATTERNS = {
  none: "none",
  solid: "solid",
  double: "double",
  dotted: "dotted",
  dashed: "dashed",
  groove: "solid",
  ridge: "solid",
  inset: "solid",
  outset: "solid",
} as const satisfies Record<CssBorderStyle, StrokePattern>;

/**
 * Every `ST_PresetLineDashVal` member, as a display-list pattern.
 *
 * DrawingML distinguishes dash length (`lgDash`) and the system patterns
 * (`sysDash`) from the plain ones; a display list has three periodic patterns,
 * so the distinction collapses to dotted or dashed. It collapses here, once,
 * rather than at whichever consumer happened to know the spelling.
 */
export const PRESET_DASH_STROKE_PATTERNS = {
  solid: "solid",
  dot: "dotted",
  dash: "dashed",
  lgDash: "dashed",
  dashDot: "dashed",
  lgDashDot: "dashed",
  lgDashDotDot: "dashed",
  sysDash: "dashed",
  sysDot: "dotted",
  sysDashDot: "dashed",
  sysDashDotDot: "dashed",
} as const satisfies Record<PresetLineDashVal, DisplayStrokePattern>;

/**
 * Every `ST_Underline` member the model carries, as the pattern one of its
 * strokes is drawn with.
 *
 * The `*Heavy` members differ from their plain counterparts in weight, which
 * `UNDERLINE_WEIGHTS` carries, not in pattern. `none` states no decoration; a
 * run whose underline says so paints nothing. `wavyDouble` is two waves, which
 * is a pattern drawn twice rather than the `double` pattern: see
 * `UNDERLINE_STROKE_COUNTS`.
 */
export const UNDERLINE_STROKE_PATTERNS = {
  none: "none",
  single: "solid",
  words: "solid",
  double: "double",
  thick: "solid",
  dotted: "dotted",
  dottedHeavy: "dotted",
  dash: "dashed",
  dashedHeavy: "dashed",
  dashLong: "dashed",
  dashLongHeavy: "dashed",
  dotDash: "dashed",
  dashDotHeavy: "dashed",
  dotDotDash: "dashed",
  dashDotDotHeavy: "dashed",
  wave: "wavy",
  wavyHeavy: "wavy",
  wavyDouble: "wavy",
} as const satisfies Record<UnderlineStyle, StrokePattern>;

/**
 * Every member's weight, as a multiple of the plain underline thickness.
 *
 * Word draws `thick` and the seven `*Heavy` members at twice the rule the font
 * asks for; the rest are the plain rule. Without this the display list strokes
 * every member at the plain thickness, so `dottedHeavy` and `dotted` come out
 * as the same line on the page and in the PDF while the editor draws them
 * apart. `none` paints nothing, so its weight is the plain one by default
 * rather than by decision.
 */
export const UNDERLINE_WEIGHTS = {
  none: PLAIN_UNDERLINE_WEIGHT,
  single: PLAIN_UNDERLINE_WEIGHT,
  words: PLAIN_UNDERLINE_WEIGHT,
  double: PLAIN_UNDERLINE_WEIGHT,
  thick: HEAVY_UNDERLINE_WEIGHT,
  dotted: PLAIN_UNDERLINE_WEIGHT,
  dottedHeavy: HEAVY_UNDERLINE_WEIGHT,
  dash: PLAIN_UNDERLINE_WEIGHT,
  dashedHeavy: HEAVY_UNDERLINE_WEIGHT,
  dashLong: PLAIN_UNDERLINE_WEIGHT,
  dashLongHeavy: HEAVY_UNDERLINE_WEIGHT,
  dotDash: PLAIN_UNDERLINE_WEIGHT,
  dashDotHeavy: HEAVY_UNDERLINE_WEIGHT,
  dotDotDash: PLAIN_UNDERLINE_WEIGHT,
  dashDotDotHeavy: HEAVY_UNDERLINE_WEIGHT,
  wave: PLAIN_UNDERLINE_WEIGHT,
  wavyHeavy: HEAVY_UNDERLINE_WEIGHT,
  wavyDouble: PLAIN_UNDERLINE_WEIGHT,
} as const satisfies Record<UnderlineStyle, number>;

/**
 * How many strokes a member emits, stacked across the underline's centre.
 *
 * `double` is one stroke: the display list's `double` pattern already carries
 * the second rule and the gap between them. `wavyDouble` is Word's two waves,
 * and no single pattern spells a doubled wave, so the builder emits the `wavy`
 * pattern twice. CSS cannot draw that at all, so the editor approximates the
 * member with two straight rules; the page and the PDF do not have to.
 */
export const UNDERLINE_STROKE_COUNTS = {
  none: 1,
  single: 1,
  words: 1,
  double: 1,
  thick: 1,
  dotted: 1,
  dottedHeavy: 1,
  dash: 1,
  dashedHeavy: 1,
  dashLong: 1,
  dashLongHeavy: 1,
  dotDash: 1,
  dashDotHeavy: 1,
  dotDotDash: 1,
  dashDotDotHeavy: 1,
  wave: 1,
  wavyHeavy: 1,
  wavyDouble: 2,
} as const satisfies Record<UnderlineStyle, number>;

/** What a border with a width but no style paints: the CSS initial. */
const DEFAULT_PATTERN: DisplayStrokePattern = "solid";

export type ResolveStrokeResult = {
  readonly stroke?: DisplayStroke;
  /** The authored colour, when it could not be parsed. Report it, do not paint. */
  readonly unresolvedColor?: string;
};

const strokeWith = (border: BorderInput, pattern: StrokePattern): ResolveStrokeResult => {
  if (pattern === "none") {
    return {};
  }

  const css = resolveCssBorderStroke(border);
  // A sub-pixel authored width comes back from the painter's helper as a
  // 1px stroke whose colour carries the fractional coverage. Undo exactly that
  // substitution: keep the authored width, keep the plain colour.
  const isSubPixelEncoded = css.color.startsWith("color-mix");
  const authoredWidth = border.width ?? css.width;
  const thicknessPx = isSubPixelEncoded ? authoredWidth : css.width;
  if (thicknessPx <= 0) {
    return {};
  }

  const colorSource = isSubPixelEncoded ? (border.color ?? "#000000") : css.color;
  const color = parseDisplayColor(colorSource);
  if (!color) {
    return { unresolvedColor: colorSource };
  }

  return { stroke: { color, thicknessPx, pattern } };
};

/**
 * Turn an authored CSS border into a paintable stroke. `undefined` stroke means
 * "draw nothing": an explicit `none` style, a non-positive width, or a colour
 * this producer cannot resolve.
 */
export const resolveBorderStroke = (border: BorderInput): ResolveStrokeResult =>
  strokeWith(
    border,
    border.style === undefined ? DEFAULT_PATTERN : CSS_BORDER_STROKE_PATTERNS[border.style],
  );

type OutlineInput = { color?: string; dash?: PresetLineDashVal; width?: number };

/** Turn a shape or text-box outline into a paintable stroke. */
export const resolveOutlineStroke = (outline: OutlineInput): ResolveStrokeResult =>
  strokeWith(
    { color: outline.color, width: outline.width },
    outline.dash === undefined ? DEFAULT_PATTERN : PRESET_DASH_STROKE_PATTERNS[outline.dash],
  );

/**
 * An authored underline's geometry. An undefined member is a `w:u` with no
 * `w:val`, which Word draws as the plain single rule.
 */
export const underlinePattern = (style: UnderlineStyle | undefined): StrokePattern =>
  style === undefined ? DEFAULT_PATTERN : UNDERLINE_STROKE_PATTERNS[style];

export const underlineWeight = (style: UnderlineStyle | undefined): number =>
  style === undefined ? PLAIN_UNDERLINE_WEIGHT : UNDERLINE_WEIGHTS[style];

export const underlineStrokeCount = (style: UnderlineStyle | undefined): number =>
  style === undefined ? 1 : UNDERLINE_STROKE_COUNTS[style];
