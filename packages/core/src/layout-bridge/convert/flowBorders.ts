/**
 * Converts document-model border specifications to layout BorderStyle.
 */

import { statesNoBorder } from "@stll/docx-core/model";
import type { BorderStyle } from "../../layout-engine/types";
import type { BorderStyleValue, ColorValue, Theme } from "../../types/document";
import { cssBorderStyle } from "../../utils/borderCss";
import { resolveColor } from "../../utils/colorResolver";
import { pointsToPixels } from "../../utils/units";

/**
 * Convert border width from eighths of a point to pixels.
 * OOXML stores border widths in eighths of a point.
 */
function borderWidthToPixels(eighthsOfPoint: number): number {
  return pointsToPixels(eighthsOfPoint / 8);
}

/** A parsed border, or the same shape read off ProseMirror attributes. */
export type ConvertibleBorder = {
  style?: BorderStyleValue;
  size?: number;
  space?: number;
  color?: ColorValue;
};

/**
 * Convert an OOXML BorderSpec to a layout-engine BorderStyle.
 * Shared by paragraph borders, cell borders, and header/footer borders.
 */
export function convertBorderSpecToLayout(
  border: ConvertibleBorder,
  theme?: Theme | null,
): BorderStyle | undefined {
  if (border.style === undefined || statesNoBorder(border.style)) {
    return undefined;
  }
  const result: BorderStyle = {
    style: cssBorderStyle(border.style),
    width: border.size === undefined ? 1 : borderWidthToPixels(border.size),
    color: border.color ? resolveColor(border.color, theme) : "#000000",
  };
  if (border.space !== undefined) {
    result.space = pointsToPixels(border.space);
  }
  return result;
}
