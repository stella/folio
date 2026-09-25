/**
 * Text-box shape helpers: inline effect-extent insets and linear gradient
 * fill resolution.
 */

import type { ImagePadding } from "@stll/docx-core/model";
import type { TextBoxBlock, TextBoxGradientFill } from "../../layout-engine/types";
import type { ShapeFill, Theme } from "../../types/document";
import { resolveColor } from "../../utils/colorResolver";

const EMUS_PER_PIXEL = 9525;

/**
 * The `wp:effectExtent` of an inline (`wp:inline`) box, in unrounded pixels.
 *
 * An inline object occupies its extent plus the effect extent on every side,
 * so the box itself sits `l`/`t` inside the space the line gives it. On a
 * `wp:anchor` the position offsets already locate the shape's own extent, so
 * the reservation only affects wrapping and is not applied here.
 */
export function inlineEffectExtentPx(
  wrapType: TextBoxBlock["wrapType"],
  extent: ImagePadding | undefined,
): TextBoxBlock["effectExtent"] {
  if ((wrapType !== undefined && wrapType !== "inline") || extent === undefined) {
    return undefined;
  }
  const px = (emu: number | undefined) => Math.max(0, (emu ?? 0) / EMUS_PER_PIXEL);
  const resolved = {
    top: px(extent.top),
    bottom: px(extent.bottom),
    left: px(extent.left),
    right: px(extent.right),
  };
  return resolved.top === 0 && resolved.bottom === 0 && resolved.left === 0 && resolved.right === 0
    ? undefined
    : resolved;
}

const GRADIENT_POSITION_SCALE = 100_000;

/**
 * A linear `a:gradFill` with its stop colors resolved through the theme. A path
 * gradient (`a:path`) is not painted: nothing here places its focus.
 */
export function resolveLinearGradientFill(
  fill: ShapeFill | undefined,
  theme: Theme | null | undefined,
): TextBoxGradientFill | undefined {
  const gradient = fill?.gradient;
  if (gradient === undefined || gradient.type !== "linear" || gradient.stops.length === 0) {
    return undefined;
  }
  const stops = gradient.stops
    .map((stop) => ({
      offset: Math.min(1, Math.max(0, stop.position / GRADIENT_POSITION_SCALE)),
      color: resolveColor(stop.color, theme),
    }))
    .toSorted((left, right) => left.offset - right.offset);
  return { angle: gradient.angle ?? 0, scaled: gradient.scaled ?? false, stops };
}
