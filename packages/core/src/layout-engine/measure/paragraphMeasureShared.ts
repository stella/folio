/**
 * Shared by the paragraph measurement modules: default metrics, the fit
 * tolerance, line state and typography shapes, run-kind guards, and run font
 * styles.
 */

import type {
  MeasuredLine,
  Run,
  TextRun,
  TabRun,
  ImageRun,
  LineBreakRun,
  FieldRun,
  MathRun,
} from "../types";
import { isFloatingImageRun } from "../types";
import type { FloatingLineSegmentZone } from "./floatingZones";
import { buildRunFontStyle, DEFAULT_FONT_FAMILY } from "./measureHelpers";
import type { FontMetrics, FontStyle } from "./measureTypes";

// Default values - match OOXML spec defaults
export const DEFAULT_FONT_SIZE = 11; // 11pt (Word 2007+ default)
export const DEFAULT_LINE_HEIGHT_MULTIPLIER = 1; // OOXML spec default: single spacing (line=240)

// Floating-point tolerance for line breaking (0.5px)
// Prevents premature line breaks due to measurement rounding
export const WIDTH_TOLERANCE = 0.5;

/**
 * Typography metrics for a line
 */
export type LineTypography = {
  ascent: number;
  descent: number;
  lineHeight: number;
};

/**
 * State tracking for line accumulation
 */
export type LineState = {
  fromRun: number;
  fromChar: number;
  toRun: number;
  toChar: number;
  width: number;
  /** Width of collapsible ASCII whitespace at the current line tail. */
  trailingWhitespaceWidth: number;
  /** Spaces the layout may compress while justifying this line. */
  regularSpaceWidth: number;
  /** Paint plan set only when final-text contraction admits a candidate. */
  justificationPaint?: NonNullable<MeasuredLine["justificationPaint"]>;
  maxFontSize: number;
  maxFontMetrics: FontMetrics | null;
  /**
   * Largest tab font on the line. Tabs are whitespace: their font sizes the
   * line only when no other run on it does (a line of nothing but tabs).
   */
  maxTabFontSize: number;
  maxTabFontMetrics: FontMetrics | null;
  /** Maximum inline image height in pixels (already in px, not points) */
  maxImageHeightPx: number;
  /** Maximum exact-height embedded-object preview on the line. */
  maxExactImageHeightPx: number;
  /** Maximum inline math height in pixels (already in px, not points) */
  maxMathHeightPx: number;
  availableWidth: number;
  /** Left offset from floating images (pixels from content left edge) */
  leftOffset: number;
  /** Right offset from floating images (pixels from content right edge) */
  rightOffset: number;
  /** Optional split segment zones from centered floating exclusions */
  segmentZones?: FloatingLineSegmentZone[];
  discretionaryHyphen?: { runIndex: number };
  renderedPageBreakBefore?: boolean;
};

/**
 * Extract FontStyle from a run that carries RunFormatting (text, tab, or
 * field). All three share the same formatting shape, so they measure the
 * same way; widening the parameter keeps tab-following measurement
 * (FieldRun page numbers, etc.) consistent with TextRun handling.
 */
export function runToFontStyle(run: TextRun | TabRun | FieldRun | MathRun): FontStyle {
  return buildRunFontStyle(run, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE);
}

/**
 * Check if a run is a text run
 */
export function isTextRun(run: Run): run is TextRun {
  return run.kind === "text";
}

/**
 * Check if a run is a tab run
 */
export function isTabRun(run: Run): run is TabRun {
  return run.kind === "tab";
}

/**
 * Check if a run is an image run
 */
export function isImageRun(run: Run): run is ImageRun {
  return run.kind === "image";
}

/**
 * Check if a run is a line break run
 */
export function isLineBreakRun(run: Run): run is LineBreakRun {
  return run.kind === "lineBreak";
}

/**
 * Check if a run is a field run
 */
export function isFieldRun(run: Run): run is FieldRun {
  return run.kind === "field";
}

/** Text a field run is measured at: its resolved value when known (so the
 *  measurer agrees with the painter), else the cached fallback. */
export function fieldMeasureText(
  run: FieldRun,
  fieldValues: ReadonlyMap<number, string> | undefined,
): string {
  const resolved = run.pmStart === undefined ? undefined : fieldValues?.get(run.pmStart);
  return resolved ?? (run.fallback || "1");
}

/**
 * Check if a run is a math equation run
 */
export function isMathRun(run: Run): run is MathRun {
  return run.kind === "math";
}

/**
 * Check if text run is empty (only whitespace or no text)
 */
export function isEmptyTextRun(run: TextRun): boolean {
  return !run.text || run.text.replace(/\u00a0/gu, " ").trim().length === 0;
}

/**
 * Sum the inline pixel widths of runs after a tab, up to (but not including)
 * the next tab or line break. Measured per-run so widths reserved match what
 * the painter draws even when trailing runs use different fonts/sizes.
 *
 * Floating/anchored images are skipped — the painter lifts them out of the
 * paragraph flow (see `isFloatingImageRun`) and `measureFollowingContentWidth`
 * in the painter already excludes them, so counting their width here would
 * desync measurer and painter on tab advance for paragraphs with a tab
 * preceding a floating image.
 */
export function isBlockLayoutImageRun(run: ImageRun): boolean {
  if (isFloatingImageRun(run)) {
    return false;
  }
  return run.wrapType === "topAndBottom" || run.displayMode === "block";
}

export function isSpaceOrTab(char: string | undefined): boolean {
  return char === " " || char === "\t";
}
