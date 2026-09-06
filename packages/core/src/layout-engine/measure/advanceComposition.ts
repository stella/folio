/**
 * The arithmetic that turns raw glyph advances into measured widths.
 *
 * Text transform, per-script font selection, letter spacing and horizontal
 * scale are the same calculation whoever supplies the advances: a canvas, a
 * parsed font, a shaper. Only the advance source differs. Keeping the
 * calculation here and parameterizing the source means a second measurement
 * backend can disagree with the canvas one about a glyph's width (which the
 * differential tests measure) but cannot disagree about where letter spacing
 * goes or which font a CJK code point takes.
 *
 * Pure: no canvas, no DOM, no caching.
 */

import {
  applyComplexScriptFormatting,
  hasComplexScriptFormatting,
} from "./complexScriptFormatting";
import { getHorizontalScaleFactor } from "../../utils/horizontalScale";
import {
  SCRIPT_CLASS,
  hasComplexScript,
  scriptClassOf,
  segmentByScript,
} from "../../utils/scriptSegments";
import type { ScriptClass } from "../../utils/scriptSegments";
import type { FontMetrics, FontStyle, RunMeasurement } from "./measureTypes";

/**
 * Raw advance of `text` in `style`, with no letter spacing, no horizontal
 * scale and no text transform applied: those are this module's job.
 */
export type GlyphAdvanceFn = (text: string, style: FontStyle) => number;

export const applyMeasurementTextTransform = (text: string, style: FontStyle): string =>
  style.textTransform === "uppercase" ? text.toLocaleUpperCase() : text;

const countCodePoints = (text: string): number => [...text].length;

/** Resolve the independent font/style slot a script segment selects. */
export const scriptStyle = (style: FontStyle, script: ScriptClass): FontStyle => {
  if (script === SCRIPT_CLASS.eastAsia) {
    if (style.eastAsiaFontFamily === undefined) {
      return style;
    }
    const result = { ...style, fontFamily: style.eastAsiaFontFamily };
    delete result.alternateFontFamily;
    if (style.eastAsiaAlternateFontFamily !== undefined) {
      result.alternateFontFamily = style.eastAsiaAlternateFontFamily;
    }
    return result;
  }
  if (script === SCRIPT_CLASS.complex) {
    return applyComplexScriptFormatting(style, style);
  }
  return style;
};

/**
 * Keep only the properties that change a glyph's advance, so a per-script
 * segment takes the single-font path: letter spacing and horizontal scale are
 * applied once over the whole string, never per segment.
 */
export const glyphAdvanceStyle = (style: FontStyle): FontStyle => {
  const result: FontStyle = {};
  if (style.fontFamily !== undefined) {
    result.fontFamily = style.fontFamily;
  }
  if (style.alternateFontFamily !== undefined) {
    result.alternateFontFamily = style.alternateFontFamily;
  }
  if (style.fontSize !== undefined) {
    result.fontSize = style.fontSize;
  }
  if (style.bold !== undefined) {
    result.bold = style.bold;
  }
  if (style.italic !== undefined) {
    result.italic = style.italic;
  }
  if (style.fontVariant !== undefined) {
    result.fontVariant = style.fontVariant;
  }
  if (style.kerning !== undefined) {
    result.kerning = style.kerning;
  }
  return result;
};

const needsPerScriptFonts = (style: FontStyle, text: string): boolean =>
  (style.eastAsiaFontFamily !== undefined && [...text].some((c) => isEastAsia(c))) ||
  (hasComplexScriptFormatting(style) && hasComplexScript(text));

const isEastAsia = (char: string): boolean => {
  // SAFETY: iterating a string yields whole code points.
  const cp = char.codePointAt(0)!;
  return scriptClassOf(cp) === SCRIPT_CLASS.eastAsia;
};

type ComposeTextWidthOptions = {
  readonly text: string;
  readonly style: FontStyle;
  readonly advanceOf: GlyphAdvanceFn;
};

/**
 * Width of a string: the advance width used for line breaking, never the
 * painted (ink) extent, which includes overhang a break decision must ignore.
 */
export const composeTextWidth = ({
  text,
  style: source,
  advanceOf,
}: ComposeTextWidthOptions): number => {
  if (!text) {
    return 0;
  }
  const style = source.forceComplexScript ? applyComplexScriptFormatting(source, source) : source;
  const transformed = applyMeasurementTextTransform(text, style);
  const letterSpacing = style.letterSpacing ?? 0;
  const horizontalScale = getHorizontalScaleFactor(style.horizontalScale);

  // Letter spacing keeps the base font for every script: CSS letter-spacing
  // does not add a gap across the per-script sibling spans the painter emits,
  // so measurement matches painting only if it stays on one font too.
  const glyphWidth =
    !source.forceComplexScript && !letterSpacing && needsPerScriptFonts(style, transformed)
      ? segmentByScript(transformed).reduce(
          (sum, segment) =>
            sum + advanceOf(segment.text, glyphAdvanceStyle(scriptStyle(style, segment.script))),
          0,
        )
      : advanceOf(transformed, style);

  const codePoints = countCodePoints(transformed);
  const spaced =
    letterSpacing && codePoints > 1 ? glyphWidth + letterSpacing * (codePoints - 1) : glyphWidth;
  return spaced * horizontalScale;
};

type ComposeRunMeasurementOptions = {
  readonly text: string;
  readonly style: FontStyle;
  readonly metrics: FontMetrics;
  readonly advanceOf: GlyphAdvanceFn;
};

/**
 * Per-code-point advances for click positioning.
 *
 * Iterates whole code points so an astral ideograph takes the East-Asian font
 * and gets a real width, while `charWidths` stays one entry per UTF-16 unit
 * (the trailing unit of a surrogate pair carries 0) so it keeps aligning with
 * ProseMirror offsets. The total differs from {@link composeTextWidth} for a
 * kerned or ligature-forming run, because each code point is measured alone.
 */
export const composeRunMeasurement = ({
  text,
  style: source,
  metrics,
  advanceOf,
}: ComposeRunMeasurementOptions): RunMeasurement => {
  const style = source.forceComplexScript ? applyComplexScriptFormatting(source, source) : source;
  if (!text) {
    return { width: 0, charWidths: [], metrics };
  }

  const perScript = !style.letterSpacing;
  const letterSpacing = style.letterSpacing ?? 0;
  const scale = getHorizontalScaleFactor(style.horizontalScale);
  const charWidths: number[] = [];
  let totalWidth = 0;
  let offset = 0;

  for (const char of text) {
    // SAFETY: iterating a string yields whole code points.
    const cp = char.codePointAt(0)!;
    const charStyle = perScript ? scriptStyle(style, scriptClassOf(cp)) : style;
    let charWidth = advanceOf(applyMeasurementTextTransform(char, style), charStyle);
    if (letterSpacing && offset + char.length < text.length) {
      charWidth += letterSpacing;
    }
    charWidth *= scale;
    charWidths.push(charWidth);
    if (char.length === 2) {
      charWidths.push(0);
    }
    totalWidth += charWidth;
    offset += char.length;
  }

  return { width: totalWidth, charWidths, metrics };
};
