/**
 * Text → one advance per code point, taken from the measure seam.
 *
 * Shared by everything that paints glyphs (body lines, list markers, tab
 * leaders, watermark text) so no caller invents a width of its own: the numbers
 * here are the numbers line breaking and pagination were decided on.
 */

import { measureRun, measureTextWidth } from "../../layout-engine/measure/measureProvider";
import {
  FONT_KERNING_MODE,
  getFontKerningMode,
} from "../../layout-engine/measure/textMeasurementPolicy";
import type { FontStyle } from "../../layout-engine/measure/measureTypes";
import { getHorizontalScaleFactor } from "../../utils/horizontalScale";
import type { DisplayRunAdjustments } from "../types";

export type Glyphs = {
  readonly text: string;
  readonly advancesPx: number[];
  readonly widthPx: number;
  /** Whether these advances were measured with kerning on. */
  readonly kerning: boolean;
  /** Whether they were measured as small capitals. */
  readonly smallCaps: boolean;
  /**
   * What went into the advances beyond the glyphs' own widths. Absent when
   * nothing did, which keeps the common run's primitive as small as it was.
   */
  readonly adjustments?: DisplayRunAdjustments;
};

export type BuildGlyphsOptions = {
  readonly text: string;
  readonly style: FontStyle;
  readonly allCaps: boolean;
  /** Justification delta applied to every compressible space. */
  readonly spaceDeltaPx: number;
  /**
   * A line-edge space run Word keeps addressable but paints with no advance
   * (the painter zeroes the split-off span's font size instead).
   */
  readonly collapsed: boolean;
};

/**
 * `measureRun` returns `charWidths` indexed by UTF-16 unit with the trailing
 * unit of a surrogate pair set to 0, because that is what ProseMirror offsets
 * need. The display list wants code points, so the pair folds into one entry.
 * An uppercase transform that expands a code point (ß → SS) keeps the source
 * advance on the first output code point and gives the rest zero, so the sum
 * still equals what the measurer decided the line on.
 */
export const buildGlyphs = (options: BuildGlyphsOptions): Glyphs => {
  const { text, style, allCaps, spaceDeltaPx, collapsed } = options;
  const adjustments = adjustmentsOf(options);
  const { charWidths } = measureRun(text, style);
  const advancesPx: number[] = [];
  let painted = "";
  let widthPx = 0;
  let unitOffset = 0;
  const shapedScale = shapedScaleOf(text, style, charWidths);

  for (const char of text) {
    const measured = (charWidths[unitOffset] ?? 0) * shapedScale;
    unitOffset += char.length;
    const advance = collapsed ? 0 : measured + (char === " " ? spaceDeltaPx : 0);
    widthPx += advance;

    let isFirst = true;
    for (const outputChar of allCaps ? char.toLocaleUpperCase() : char) {
      painted += outputChar;
      advancesPx.push(isFirst ? advance : 0);
      isFirst = false;
    }
  }

  return {
    text: painted,
    advancesPx,
    widthPx,
    kerning: getFontKerningMode(style) === FONT_KERNING_MODE.enabled,
    smallCaps: style.fontVariant === "small-caps",
    ...(adjustments === undefined ? {} : { adjustments }),
  };
};

/**
 * What every per-character width has to be multiplied by for the run to occupy
 * the width the line was fitted at.
 *
 * The two numbers come from the same measurer and differ by what shaping does
 * across a boundary: a pair that kerns or ligates is narrower in the string
 * than the two characters are apart. Line breaking used the string's width, so
 * that is the run's width, and advances that summed to anything else would
 * declare an extent no backend paints.
 *
 * Spread over every character rather than dropped on the last, because the
 * difference is not one boundary's: concentrating it would move the run's final
 * glyph by the whole of it.
 */
const shapedScaleOf = (text: string, style: FontStyle, charWidths: readonly number[]): number => {
  const perCharacterPx = charWidths.reduce((total, width) => total + width, 0);
  if (perCharacterPx <= 0) {
    return 1;
  }
  return measureTextWidth(text, style) / perCharacterPx;
};

/**
 * The three adjustments named, or nothing when none applies.
 *
 * A collapsed run carries no advances at all, so it has nothing to reapply
 * either: naming a spacing on a run painted at zero width would have a backend
 * widen what the line was fitted without.
 */
const adjustmentsOf = ({
  style,
  spaceDeltaPx,
  collapsed,
}: BuildGlyphsOptions): DisplayRunAdjustments | undefined => {
  const horizontalScale = getHorizontalScaleFactor(style.horizontalScale);
  // Measured letter spacing is pre-scale, as CSS letter-spacing is; the advances
  // carry it scaled, so it is reported the same way.
  const letterSpacingPx = collapsed ? 0 : (style.letterSpacing ?? 0) * horizontalScale;
  const wordSpacingPx = collapsed ? 0 : spaceDeltaPx;
  if (letterSpacingPx === 0 && wordSpacingPx === 0 && horizontalScale === 1) {
    return undefined;
  }
  return { letterSpacingPx, horizontalScale, wordSpacingPx };
};

/**
 * Everything a measured run contributes to a `glyphRun`.
 *
 * One helper rather than three fields at each emission site: what went into the
 * advances has to travel with them, and a call site that carried the advances
 * and forgot the adjustments would silently paint a run the line was not fitted
 * with.
 */
export const glyphRunText = (glyphs: Glyphs) => ({
  text: glyphs.text,
  advancesPx: glyphs.advancesPx,
  kerning: glyphs.kerning,
  smallCaps: glyphs.smallCaps,
  ...(glyphs.adjustments === undefined ? {} : { adjustments: glyphs.adjustments }),
});
