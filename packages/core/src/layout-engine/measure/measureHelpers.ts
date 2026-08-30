/**
 * Pure measurement helpers — no canvas, no DOM, no worker.
 *
 * Unit conversions, font-string construction, run → FontStyle mapping, and
 * char-offset geometry. The layout engine imports these directly so it never
 * transitively pulls the canvas measurement backend; the canvas implementation
 * (`measureContainer.ts`) consumes them too.
 * Script sizing: eigenpal/docx-editor@585413d0 (Apache-2.0), modified for Folio.
 */

import { resolveFontFamily } from "../../utils/fontResolver";
import { DOCX_BOLD_FONT_WEIGHT } from "../../utils/fontWeights";
import { normalizeHorizontalScalePercent } from "../../utils/horizontalScale";
import type { RunFormatting } from "../types";
import type { FontStyle } from "./measureTypes";
import { FONT_KERNING_MODE, getRunFontKerningMode } from "./textMeasurementPolicy";

// Constants for OOXML unit conversions
const TWIPS_PER_INCH = 1440;
const PX_PER_INCH = 96; // Standard CSS/DOM DPI
const TWIPS_PER_PX = TWIPS_PER_INCH / PX_PER_INCH; // 15 twips per pixel

// Default typography values
export const DEFAULT_FONT_SIZE = 11; // 11pt (Word 2007+ default)
export const DEFAULT_FONT_FAMILY = "Calibri";
export const DOCX_SCRIPT_FONT_SCALE = 0.75;

/**
 * Build a measurement `FontStyle` from a run's formatting. Single source of
 * truth for run → FontStyle so every measurement path (layout line-breaking,
 * click-to-position, selection rects) carries the same fields — notably
 * `eastAsiaFontFamily`, which must reach the measurer for CJK click/selection
 * offsets to match what was wrapped and painted. Callers pass their own family
 * and size fallbacks for runs that declare neither.
 */
export function buildRunFontStyle(
  run: RunFormatting,
  fallbackFontFamily: string,
  fallbackFontSize: number,
): FontStyle {
  const baseFontSize = run.fontSize ?? fallbackFontSize;
  const fontSize =
    run.superscript || run.subscript ? baseFontSize * DOCX_SCRIPT_FONT_SCALE : baseFontSize;
  const horizontalScale = normalizeHorizontalScalePercent(run.horizontalScale);
  return {
    fontFamily: run.fontFamily ?? fallbackFontFamily,
    ...(run.alternateFontFamily !== undefined
      ? { alternateFontFamily: run.alternateFontFamily }
      : {}),
    ...(run.eastAsiaFontFamily !== undefined ? { eastAsiaFontFamily: run.eastAsiaFontFamily } : {}),
    ...(run.eastAsiaAlternateFontFamily !== undefined
      ? { eastAsiaAlternateFontFamily: run.eastAsiaAlternateFontFamily }
      : {}),
    ...(run.complexScriptFontFamily !== undefined
      ? { complexScriptFontFamily: run.complexScriptFontFamily }
      : {}),
    ...(run.complexScriptAlternateFontFamily !== undefined
      ? { complexScriptAlternateFontFamily: run.complexScriptAlternateFontFamily }
      : {}),
    ...(run.complexScriptFontSize !== undefined
      ? { complexScriptFontSize: run.complexScriptFontSize }
      : {}),
    ...(run.complexScriptBold !== undefined ? { complexScriptBold: run.complexScriptBold } : {}),
    ...(run.complexScriptItalic !== undefined
      ? { complexScriptItalic: run.complexScriptItalic }
      : {}),
    ...(run.forceComplexScript !== undefined ? { forceComplexScript: run.forceComplexScript } : {}),
    fontSize,
    ...(run.bold !== undefined ? { bold: run.bold } : {}),
    ...(run.italic !== undefined ? { italic: run.italic } : {}),
    ...(run.letterSpacing !== undefined ? { letterSpacing: run.letterSpacing } : {}),
    ...(run.allCaps ? { textTransform: "uppercase" as const } : {}),
    ...(run.smallCaps ? { fontVariant: "small-caps" as const } : {}),
    ...(horizontalScale !== undefined ? { horizontalScale } : {}),
    kerning: getRunFontKerningMode(run, fallbackFontSize) === FONT_KERNING_MODE.enabled,
  };
}

/** Cached resolved font data (CSS fallback + single-line ratio) */
type ResolvedFontCache = {
  cssFallback: string;
  singleLineRatio: number;
};

/** Cache for resolved font data */
const fontResolvedCache = new Map<string, ResolvedFontCache>();

/**
 * Get the resolved font data for a font family, with caching.
 */
export function getResolvedData(
  fontFamily: string,
  alternateFontFamily?: string,
): ResolvedFontCache {
  const cacheKey = `${fontFamily}\u0000${alternateFontFamily ?? ""}`;
  let cached = fontResolvedCache.get(cacheKey);
  if (cached === undefined) {
    const resolved = resolveFontFamily(fontFamily, alternateFontFamily);
    cached = {
      cssFallback: resolved.cssFallback,
      singleLineRatio: resolved.singleLineRatio,
    };
    fontResolvedCache.set(cacheKey, cached);
  }
  return cached;
}

/**
 * Get the CSS fallback string for a font family, with caching.
 */
function getResolvedFallback(fontFamily: string, alternateFontFamily?: string): string {
  return getResolvedData(fontFamily, alternateFontFamily).cssFallback;
}

/**
 * Clear the resolved-font-family cache. `resolveFontFamily`'s output for a
 * given key can change even when the raw font names don't — e.g. when the
 * active document's embedded-font family map changes (see
 * `utils/fontResolver.ts`'s `setEmbeddedFontFamilyMap`) — so this must be
 * cleared whenever that happens, not just on font-name changes. Folded into
 * `clearAllCaches` (`./cache.ts`) so callers get it for free.
 */
export function clearFontResolvedCache(): void {
  fontResolvedCache.clear();
}

/**
 * Build a CSS font string from styling properties
 *
 * Font sizes are in points and need to be converted to pixels for canvas.
 * 1pt = 96/72 px ≈ 1.333px at standard web DPI.
 *
 * Uses the font resolver to get category-appropriate fallback stacks
 * (serif fonts get serif fallbacks, sans-serif get sans-serif, etc.)
 * matching the same stacks used in rendering for consistent measurements.
 *
 * @example
 * buildFontString({ fontFamily: "Arial", fontSize: 12, bold: true })
 * // Returns: "700 16px Arial, Arimo, Helvetica, sans-serif" (12pt = 16px)
 */
export function buildFontString(style: FontStyle): string {
  const parts: string[] = [];

  if (style.italic) {
    parts.push("italic");
  }
  if (style.fontVariant) {
    parts.push(style.fontVariant);
  }
  if (style.bold) {
    parts.push(DOCX_BOLD_FONT_WEIGHT);
  }

  // Convert points to pixels for canvas measurement
  const fontSizePt = style.fontSize ?? DEFAULT_FONT_SIZE;
  const fontSizePx = ptToPx(fontSizePt);
  parts.push(`${fontSizePx}px`);

  // Use the font resolver for category-appropriate fallback stacks
  const fontFamily = style.fontFamily ?? DEFAULT_FONT_FAMILY;
  parts.push(getResolvedFallback(fontFamily, style.alternateFontFamily));

  return parts.join(" ");
}

/**
 * Find the character offset at a given X position within a text run
 *
 * @param x - X position relative to run start
 * @param charWidths - Per-character widths from measureRun
 * @returns Character offset (0-based index)
 */
export function findCharacterAtX(x: number, charWidths: number[]): number {
  if (charWidths.length === 0) {
    return 0;
  }
  if (x <= 0) {
    return 0;
  }

  let accumulatedWidth = 0;

  for (let i = 0; i < charWidths.length; i++) {
    // SAFETY: i < charWidths.length in for loop
    const charWidth = charWidths[i]!;
    const charMidpoint = accumulatedWidth + charWidth / 2;

    // If x is before the midpoint, the cursor is at this character
    if (x <= charMidpoint) {
      return i;
    }

    accumulatedWidth += charWidth;
  }

  // X is past all characters, return position after last character
  return charWidths.length;
}

/**
 * Get the X position of a character offset within a text run
 *
 * @param offset - Character offset (0-based index)
 * @param charWidths - Per-character widths from measureRun
 * @returns X position in pixels
 */
export function getXForCharacter(offset: number, charWidths: number[]): number {
  if (offset <= 0 || charWidths.length === 0) {
    return 0;
  }

  const clampedOffset = Math.min(offset, charWidths.length);
  let x = 0;

  for (let i = 0; i < clampedOffset; i++) {
    // SAFETY: i < clampedOffset <= charWidths.length
    x += charWidths[i]!;
  }

  return x;
}

// Unit conversion utilities

/**
 * Convert twips to pixels
 */
export function twipsToPx(twips: number): number {
  return twips / TWIPS_PER_PX;
}

/**
 * Convert pixels to twips
 */
export function pxToTwips(px: number): number {
  return Math.round(px * TWIPS_PER_PX);
}

/**
 * Convert points to pixels
 */
export function ptToPx(pt: number): number {
  return (pt * PX_PER_INCH) / 72;
}

/**
 * Convert pixels to points
 */
export function pxToPt(px: number): number {
  return (px * 72) / PX_PER_INCH;
}

/**
 * Convert OOXML half-points to pixels
 * OOXML font sizes are in half-points (24 = 12pt)
 */
export function halfPtToPx(halfPt: number): number {
  return ptToPx(halfPt / 2);
}

/**
 * Convert pixels to OOXML half-points
 */
export function pxToHalfPt(px: number): number {
  return pxToPt(px) * 2;
}
