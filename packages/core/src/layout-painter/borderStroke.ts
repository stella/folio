import type { ParagraphBorders } from "../layout-engine/types";
import { pointsToPixels } from "../utils/units";

/**
 * A stroke to paint, in whichever vocabulary its producer speaks.
 *
 * `style` is plain `string` on purpose: a paragraph or table border arrives as
 * a CSS `border-style`, an image border as a CSS keyword the host supplied, and
 * a shape outline as a DrawingML `ST_PresetLineDashVal` (`sysDash`, `lgDashDot`
 * …). This helper only forwards it; `strokes.ts` is where the vocabularies are
 * reconciled.
 */
type BorderStroke = {
  color?: string;
  style?: string;
  width?: number;
};

type CssBorderStroke = {
  color: string;
  style: string;
  width: number;
};

const DEFAULT_BORDER_COLOR = "#000000";
const DEFAULT_BORDER_STYLE = "solid";
const CSS_HAIRLINE_WIDTH = 1;
const PARAGRAPH_RULE_ENDPOINT_OUTSET = pointsToPixels(1.5);

/**
 * Preserve the optical weight of authored subpixel borders. CSS quantizes a
 * fractional border width to one device pixel, so its color carries the
 * fractional coverage while layout keeps the authored width.
 */
export const resolveCssBorderStroke = (border: BorderStroke): CssBorderStroke => {
  const authoredWidth = border.width ?? CSS_HAIRLINE_WIDTH;
  if (authoredWidth <= 0 || authoredWidth >= CSS_HAIRLINE_WIDTH) {
    return {
      color: border.color ?? DEFAULT_BORDER_COLOR,
      style: border.style ?? DEFAULT_BORDER_STYLE,
      width: Math.max(CSS_HAIRLINE_WIDTH, authoredWidth),
    };
  }

  const coveragePercent = Number((authoredWidth * 100).toFixed(4));
  const color = border.color ?? DEFAULT_BORDER_COLOR;
  return {
    color: `color-mix(in srgb, ${color} ${coveragePercent}%, transparent)`,
    style: border.style ?? DEFAULT_BORDER_STYLE,
    width: CSS_HAIRLINE_WIDTH,
  };
};

export const borderStrokeToCss = (border: BorderStroke): string => {
  const stroke = resolveCssBorderStroke(border);
  return `${stroke.width}px ${stroke.style} ${stroke.color}`;
};

type ParagraphBorderHorizontalOutsets = {
  left: number;
  right: number;
};

/** Resolve the outer horizontal edges of a paragraph border box. */
export const resolveParagraphBorderHorizontalOutsets = (
  borders: ParagraphBorders,
  hasHorizontalRule: boolean,
): ParagraphBorderHorizontalOutsets => {
  const ruleOutset = hasHorizontalRule ? PARAGRAPH_RULE_ENDPOINT_OUTSET : 0;
  return {
    left: borders.left ? (borders.left.space ?? 0) + (borders.left.width ?? 0) : ruleOutset,
    right: borders.right ? (borders.right.space ?? 0) + (borders.right.width ?? 0) : ruleOutset,
  };
};
