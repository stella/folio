/**
 * Border and rule strokes.
 *
 * One deliberate divergence from `layout-painter/borderStroke.ts`: that helper
 * encodes a sub-pixel border width as a `color-mix` alpha because CSS quantizes
 * a fractional border to one device pixel. A display list has no device pixel,
 * and a PDF strokes 0.4pt natively, so the authored fractional width and the
 * plain colour survive here. Everything else (colour/style/width defaults) is
 * taken from the painter's helper so the two cannot drift.
 */

import { resolveCssBorderStroke } from "../../layout-painter/borderStroke";
import type { BorderStyle } from "../../layout-engine/types";
import type { DisplayStroke, DisplayStrokePattern } from "../types";
import { parseDisplayColor } from "./colors";

type BorderInput = Pick<BorderStyle, "color" | "style" | "width">;

/**
 * CSS `border-style` and the raw OOXML `w:val` names both reach
 * `BorderStyle.style` (the bridge maps most, `renderParagraph` maps the rest),
 * so both vocabularies are resolved here. Everything decorative that folio does
 * not draw degrades to a plain line, matching how Word degrades on a platform
 * without the specialised glyphs.
 */
const STROKE_PATTERN_BY_STYLE: Record<string, DisplayStrokePattern | "none"> = {
  none: "none",
  nil: "none",
  hidden: "none",
  solid: "solid",
  single: "solid",
  thick: "solid",
  inset: "solid",
  outset: "solid",
  ridge: "solid",
  groove: "solid",
  double: "double",
  dotted: "dotted",
  dotdash: "dashed",
  dotdotdash: "dashed",
  dashed: "dashed",
  dash: "dashed",
  dashsmallgap: "dashed",
  dashlong: "dashed",
  wavy: "wavy",
  wave: "wavy",
  wavydouble: "double",
};

const DEFAULT_PATTERN: DisplayStrokePattern = "solid";

const strokePatternForStyle = (style: string | undefined): DisplayStrokePattern | "none" => {
  if (style === undefined) {
    return DEFAULT_PATTERN;
  }
  return STROKE_PATTERN_BY_STYLE[style.trim().toLowerCase()] ?? DEFAULT_PATTERN;
};

export type ResolveStrokeResult = {
  readonly stroke?: DisplayStroke;
  /** The authored colour, when it could not be parsed. Report it, do not paint. */
  readonly unresolvedColor?: string;
};

/**
 * Turn an authored border into a paintable stroke. `undefined` stroke means
 * "draw nothing": an explicit `none`/`nil` style, a non-positive width, or a
 * colour this producer cannot resolve.
 */
export const resolveBorderStroke = (border: BorderInput): ResolveStrokeResult => {
  const pattern = strokePatternForStyle(border.style);
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

/** CSS `text-decoration-style` → the display list's stroke patterns. */
export const decorationPatternForStyle = (style: string | undefined): DisplayStrokePattern => {
  if (style === undefined) {
    return DEFAULT_PATTERN;
  }
  const pattern = STROKE_PATTERN_BY_STYLE[style.trim().toLowerCase()];
  return pattern === undefined || pattern === "none" ? DEFAULT_PATTERN : pattern;
};
