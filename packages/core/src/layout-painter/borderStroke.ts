import type { ParagraphBorders } from "../layout-engine/types";
import type { CssBorderStyle } from "../utils/borderCss";
import { pointsToPixels } from "../utils/units";

/**
 * A CSS border to paint.
 *
 * `style` is the CSS keyword, never an `ST_Border` member and never a
 * DrawingML dash: an outline's dash is translated by `cssBorderStyleForDash`
 * before it reaches a border shorthand, because `border: 1px sysDash #000`
 * invalidates the whole declaration.
 */
type BorderStroke = {
  color?: string | undefined;
  style?: CssBorderStyle | undefined;
  width?: number | undefined;
};

type CssBorderStroke = {
  color: string;
  style: CssBorderStyle;
  width: number;
};

const DEFAULT_BORDER_COLOR = "#000000";
const DEFAULT_BORDER_STYLE = "solid" as const satisfies CssBorderStyle;
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
