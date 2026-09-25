/**
 * Line height and typography metrics: font-driven line heights, East Asian
 * and complex-script line height styles, empty-paragraph metrics, and
 * document-grid line pitch.
 */

import {
  CJK_FALLBACK_FONT_FAMILY,
  DEFAULT_SINGLE_LINE_RATIO,
  isCjkFont,
} from "../../utils/fontResolver";
import { hasCjk, hasComplexScript, hasEastAsiaSlotText } from "../../utils/scriptSegments";
import type { TextRun, ParagraphSpacing, ParagraphAttrs } from "../types";
import {
  applyComplexScriptFormatting,
  hasComplexScriptFormatting,
} from "./complexScriptFormatting";
import { DEFAULT_FONT_FAMILY, ptToPx } from "./measureHelpers";
import { getFontMetrics } from "./measureProvider";
import type { FontMetrics, FontStyle } from "./measureTypes";
import { DEFAULT_FONT_SIZE, DEFAULT_LINE_HEIGHT_MULTIPLIER } from "./paragraphMeasureShared";
import type { LineTypography } from "./paragraphMeasureShared";

/**
 * Line-HEIGHT style for a text run. Word derives a CJK line's height from an
 * East-Asian face, not the run's ascii font: real documents routinely put CJK
 * glyphs in runs whose ascii/eastAsia fonts are Latin (e.g. `w:eastAsia`
 * "Century"), and Word font-links those glyphs to the default East-Asian face
 * and uses its taller single-line height (≈1.303 vs the ≈1.15 Latin default).
 *
 * When the run holds CJK text, swap `fontFamily` to the face whose
 * `singleLineRatio` Word would actually use: the run's `w:eastAsia` font when
 * it is a real CJK face, the run's own ascii font when THAT is a CJK face, or
 * `CJK_FALLBACK_FONT_FAMILY` otherwise. Non-CJK runs return `baseStyle`
 * unchanged, so Latin-only documents are unaffected by construction.
 *
 * The result feeds `updateMaxFont` (line height) ONLY — width measurement
 * keeps the base style so wrapping is untouched.
 */
export function cjkLineHeightStyle(run: TextRun, baseStyle: FontStyle): FontStyle {
  if (!run.text) {
    return baseStyle;
  }
  const eastAsia = baseStyle.eastAsiaFontFamily;
  const eastAsiaFace = eastAsia !== undefined && isCjkFont(eastAsia);
  // Symbols a `w:hint="eastAsia"` moves paint with the run's East Asian face,
  // so they size the line only when there is one.
  if (
    !hasCjk(run.text) &&
    !(eastAsiaFace && hasEastAsiaSlotText(run.text, baseStyle.eastAsiaHint))
  ) {
    return baseStyle;
  }
  if (eastAsia !== undefined && eastAsiaFace) {
    const result = { ...baseStyle, fontFamily: eastAsia };
    delete result.alternateFontFamily;
    if (baseStyle.eastAsiaAlternateFontFamily !== undefined) {
      result.alternateFontFamily = baseStyle.eastAsiaAlternateFontFamily;
    }
    return result;
  }
  if (isCjkFont(baseStyle.fontFamily ?? DEFAULT_FONT_FAMILY)) {
    return baseStyle;
  }
  return { ...baseStyle, fontFamily: CJK_FALLBACK_FONT_FAMILY };
}

/** Pick the dominant typography metrics for a run containing complex script. */
export function complexScriptLineHeightStyle(run: TextRun, baseStyle: FontStyle): FontStyle {
  if (
    !hasComplexScriptFormatting(run) ||
    (!run.forceComplexScript && !hasComplexScript(run.text))
  ) {
    return baseStyle;
  }
  const complexStyle = applyComplexScriptFormatting(baseStyle, run);
  return (complexStyle.fontSize ?? DEFAULT_FONT_SIZE) >= (baseStyle.fontSize ?? DEFAULT_FONT_SIZE)
    ? complexStyle
    : baseStyle;
}

/**
 * Calculate typography metrics from font size and spacing settings
 *
 * @param fontSize - Font size in points
 * @param spacing - Paragraph spacing settings
 * @param metrics - Pre-calculated font metrics (in pixels)
 */
export function calculateTypographyMetrics(
  fontSize: number,
  spacing?: ParagraphSpacing,
  metrics?: FontMetrics | null,
): LineTypography {
  // Use provided metrics or calculate from font size
  // When calculating from fontSize (points), convert to pixels first
  const fontSizePx = ptToPx(fontSize);
  const ascent = metrics?.ascent ?? fontSizePx * 0.8;
  const descent = metrics?.descent ?? fontSizePx * 0.2;

  // Apply line spacing rules
  //
  // OOXML lineRule="auto" multipliers (w:line in 240ths):
  //   line=240 → 1.0x (single), line=276 → 1.15x (Word default), line=480 → 2.0x
  //
  // The multiplier base is the font's "single line" height per OOXML spec (§17.3.1.33):
  //   singleLine = (usWinAscent + usWinDescent) / unitsPerEm × fontSizePx
  // This ratio is font-specific (1.07–1.27 for common fonts). We use a hardcoded
  // lookup table of OS/2 metrics since Canvas fontBoundingBox is unreliable
  // cross-platform (Mac uses hhea, not usWin) and Google Font substitutes
  // report different metrics than the original fonts.
  const ratio = metrics?.singleLineRatio ?? DEFAULT_SINGLE_LINE_RATIO;
  const singleLineBase = fontSizePx * ratio;

  let lineHeight: number;

  if (spacing?.lineRule === "exact" && spacing.line !== undefined) {
    // Exact: use specified height exactly
    lineHeight = spacing.line;
  } else if (spacing?.lineRule === "atLeast" && spacing.line !== undefined) {
    // At least: use specified height or natural height, whichever is larger
    const defaultHeight = singleLineBase * DEFAULT_LINE_HEIGHT_MULTIPLIER;
    lineHeight = Math.max(spacing.line, defaultHeight);
  } else if (spacing?.line !== undefined && spacing.lineUnit === "multiplier") {
    // Multiplier applied to font's single-line height
    lineHeight = singleLineBase * spacing.line;
  } else if (spacing?.line !== undefined && spacing.lineUnit === "px") {
    // Pixel value
    lineHeight = spacing.line;
  } else {
    // No explicit spacing — OOXML spec default is line=240 (1.0x = single spacing).
    // Documents wanting 1.15x set w:line=276 explicitly in styles, which flows
    // through the multiplier branch above. This fallback is for paragraphs with
    // no style and no direct formatting.
    lineHeight = singleLineBase * DEFAULT_LINE_HEIGHT_MULTIPLIER;
  }

  return { ascent, descent, lineHeight };
}

/**
 * Word's "single line spacing" floor (≈ 1.15×) applied to empty paragraphs
 * with `auto`/`atLeast` line rules. Without this, narrow-metric fonts
 * (Arial Narrow, OS/2 ratio ≈ 1.117) collapse empty rows visibly tighter
 * than Word renders them. See eigenpal #391/#394.
 */
const WORD_SINGLE_LINE_FLOOR = 1.15;

/**
 * Calculate metrics for an empty paragraph.
 *
 * Word renders an empty paragraph as a single readable line — its line
 * height never collapses below 1.15 × font size, even when the doc
 * explicitly writes `<w:line w:val="240"/>` (1.0×). The floor is scoped to
 * `auto`/`atLeast` line rules; `exact` means exact (per OOXML §17.3.1.33)
 * and stays untouched.
 */
export function calculateEmptyParagraphMetrics(
  fontSize: number,
  spacing?: ParagraphSpacing,
  fontFamily?: string,
  attrs?: ParagraphAttrs,
): LineTypography {
  const metrics = getFontMetrics({
    fontSize,
    fontFamily: fontFamily ?? DEFAULT_FONT_FAMILY,
    ...(attrs?.defaultAlternateFontFamily !== undefined
      ? { alternateFontFamily: attrs.defaultAlternateFontFamily }
      : {}),
  });
  const result = calculateTypographyMetrics(fontSize, spacing, metrics);

  const lineRule = spacing?.lineRule ?? "auto";
  if (lineRule === "auto" || lineRule === "atLeast") {
    const fontSizePx = ptToPx(fontSize);
    const floor = fontSizePx * WORD_SINGLE_LINE_FLOOR;
    if (result.lineHeight < floor) {
      return applyDocumentGrid({ ...result, lineHeight: floor }, attrs, spacing);
    }
  }
  return applyDocumentGrid(result, attrs, spacing);
}

export function applyDocumentGrid(
  typography: LineTypography,
  attrs: ParagraphAttrs | undefined,
  spacing: ParagraphSpacing | undefined,
): LineTypography {
  const pitch = attrs?.documentGridLinePitch;
  if (pitch === undefined || pitch <= 0 || attrs?.snapToGrid === false) {
    return typography;
  }
  if (spacing?.lineRule === "exact") {
    return typography;
  }
  const lineHeight = Math.ceil(typography.lineHeight / pitch - 1e-9) * pitch;
  return lineHeight === typography.lineHeight ? typography : { ...typography, lineHeight };
}
