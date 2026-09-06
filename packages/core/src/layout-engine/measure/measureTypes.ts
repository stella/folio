/**
 * Measurement data shapes.
 *
 * Pure type declarations shared by the provider seam (`measureProvider.ts`),
 * the pure helpers (`measureHelpers.ts`), and the canvas implementation
 * (`measureContainer.ts`). Types only — this module has no runtime and never
 * pulls canvas/DOM into an importer's graph.
 */

/**
 * Font styling properties for measurement
 */
export type FontStyle = {
  fontFamily?: string;
  alternateFontFamily?: string;
  /**
   * East-Asian font for CJK code points. When set, `measureTextWidth` /
   * `measureRun` measure CJK code points with this font and the rest with
   * `fontFamily`, matching the painter's per-script span split so wrapping
   * and click positioning stay in sync.
   */
  eastAsiaFontFamily?: string;
  eastAsiaAlternateFontFamily?: string;
  /**
   * Complex-script font for Arabic, Hebrew, Indic and South-East Asian code
   * points. Same contract as `eastAsiaFontFamily`, over a different slot.
   */
  complexScriptFontFamily?: string;
  complexScriptAlternateFontFamily?: string;
  complexScriptFontSize?: number;
  complexScriptBold?: boolean;
  complexScriptItalic?: boolean;
  forceComplexScript?: boolean;
  fontSize?: number; // in points
  bold?: boolean;
  italic?: boolean;
  letterSpacing?: number; // in pixels
  textTransform?: "uppercase";
  fontVariant?: "small-caps";
  horizontalScale?: number;
  /** Enable pair kerning for runs whose authored threshold is met. */
  kerning?: boolean;
};

/**
 * Typography metrics for a font
 */
export type FontMetrics = {
  fontSize: number;
  /** Ink extent above the baseline: how far the glyphs actually reach. */
  ascent: number;
  /** Ink extent below the baseline. */
  descent: number;
  /**
   * Extent of the font's own box (hhea/OS-2), which is wider than the ink and
   * is what a browser builds an inline content area from. A painter that
   * positions a text box from the ink extents instead puts every baseline out
   * by half the difference between the two, per face, with nothing to
   * attribute it to. Carried beside the ink extents rather than replacing them
   * because line boxes are sized from the ink.
   */
  fontBoxAscent: number;
  fontBoxDescent: number;
  lineHeight: number;
  fontFamily: string;
  /** OS/2 single-line ratio for OOXML line spacing calculation */
  singleLineRatio: number;
};

/**
 * Result of measuring a text string
 */
export type TextMeasurement = {
  width: number;
  height: number;
  ascent: number;
  descent: number;
};

/**
 * Result of measuring a run of text
 */
export type RunMeasurement = {
  width: number;
  charWidths: number[]; // Width of each character for click positioning
  metrics: FontMetrics;
};
