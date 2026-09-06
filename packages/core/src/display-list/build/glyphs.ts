/**
 * Text → one advance per code point, taken from the measure seam.
 *
 * Shared by everything that paints glyphs (body lines, list markers, tab
 * leaders, watermark text) so no caller invents a width of its own: the numbers
 * here are the numbers line breaking and pagination were decided on.
 */

import { measureRun } from "../../layout-engine/measure/measureProvider";
import type { FontStyle } from "../../layout-engine/measure/measureTypes";

export type Glyphs = {
  readonly text: string;
  readonly advancesPx: number[];
  readonly widthPx: number;
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
export const buildGlyphs = ({
  text,
  style,
  allCaps,
  spaceDeltaPx,
  collapsed,
}: BuildGlyphsOptions): Glyphs => {
  const { charWidths } = measureRun(text, style);
  const advancesPx: number[] = [];
  let painted = "";
  let widthPx = 0;
  let unitOffset = 0;

  for (const char of text) {
    const measured = charWidths[unitOffset] ?? 0;
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

  return { text: painted, advancesPx, widthPx };
};
