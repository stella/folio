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
 * What one `ST_Underline` member paints: the pattern each of its strokes is
 * drawn with, the weight of that stroke as a multiple of the plain underline
 * thickness, and how many strokes stack across the underline's centre.
 */
type UnderlineStrokes = {
  readonly pattern: StrokePattern;
  readonly weight: number;
  readonly strokes: 1 | 2;
};

/**
 * Every `ST_Underline` member the model carries, one row per member.
 *
 * The three fields are one decision and belong on one line: `wavyDouble` is
 * Word's two waves, which is the `wavy` pattern drawn twice rather than the
 * `double` pattern, so a reader who has the pattern without the count has half
 * the member. `double` is one stroke, because the display list's `double`
 * pattern already carries the second rule and the gap between them. CSS can
 * draw neither two waves nor `words`, so the editor approximates them; the page
 * and the PDF do not have to.
 *
 * Word draws `thick` and the seven `*Heavy` members at twice the rule the font
 * asks for, and the rest at the plain rule. Without the weight the display list
 * strokes every member at the plain thickness, so `dottedHeavy` and `dotted`
 * come out as the same line on the page and in the PDF while the editor draws
 * them apart. The DOM spells this same weight as a
 * `text-decoration-thickness` length (`utils/formatToStyle.ts`), so which
 * members are heavy is decided here and nowhere else.
 *
 * `none` states no decoration: a run whose underline says so paints nothing,
 * and its weight and stroke count are the plain ones by default rather than by
 * decision.
 */
export const UNDERLINE_STROKES = {
  none: { pattern: "none", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 1 },
  single: { pattern: "solid", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 1 },
  words: { pattern: "solid", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 1 },
  double: { pattern: "double", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 1 },
  thick: { pattern: "solid", weight: HEAVY_UNDERLINE_WEIGHT, strokes: 1 },
  dotted: { pattern: "dotted", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 1 },
  dottedHeavy: { pattern: "dotted", weight: HEAVY_UNDERLINE_WEIGHT, strokes: 1 },
  dash: { pattern: "dashed", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 1 },
  dashedHeavy: { pattern: "dashed", weight: HEAVY_UNDERLINE_WEIGHT, strokes: 1 },
  dashLong: { pattern: "dashed", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 1 },
  dashLongHeavy: { pattern: "dashed", weight: HEAVY_UNDERLINE_WEIGHT, strokes: 1 },
  dotDash: { pattern: "dashed", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 1 },
  dashDotHeavy: { pattern: "dashed", weight: HEAVY_UNDERLINE_WEIGHT, strokes: 1 },
  dotDotDash: { pattern: "dashed", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 1 },
  dashDotDotHeavy: { pattern: "dashed", weight: HEAVY_UNDERLINE_WEIGHT, strokes: 1 },
  wave: { pattern: "wavy", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 1 },
  wavyHeavy: { pattern: "wavy", weight: HEAVY_UNDERLINE_WEIGHT, strokes: 1 },
  wavyDouble: { pattern: "wavy", weight: PLAIN_UNDERLINE_WEIGHT, strokes: 2 },
} as const satisfies Record<UnderlineStyle, UnderlineStrokes>;

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
 * An authored underline's row. An undefined member is a `w:u` with no `w:val`,
 * which Word draws as the plain single rule, so it resolves to that member's
 * row rather than to a default per field.
 */
const underlineStrokes = (style: UnderlineStyle | undefined): UnderlineStrokes =>
  style === undefined ? UNDERLINE_STROKES.single : UNDERLINE_STROKES[style];

export const underlinePattern = (style: UnderlineStyle | undefined): StrokePattern =>
  underlineStrokes(style).pattern;

export const underlineWeight = (style: UnderlineStyle | undefined): number =>
  underlineStrokes(style).weight;

export const underlineStrokeCount = (style: UnderlineStyle | undefined): number =>
  underlineStrokes(style).strokes;
