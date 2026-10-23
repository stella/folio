/**
 * Swappable text-measurement seam.
 *
 * Holds the provider interface, the active-provider registry, and the four
 * public delegators the layout engine measures through. This module is pure:
 * it never imports the canvas backend, so the engine's import graph stays
 * canvas-free. The default provider throws — a composition root must install a
 * concrete backend first (the React editor and the bun test preload call
 * `installCanvasMeasureProvider()`).
 */

import { panic } from "better-result";

import {
  automaticGapCodeUnitAdvances,
  automaticTextSpacingGaps,
} from "./automaticTextSpacing";
import type {
  FontMetrics,
  FontStyle,
  RunMeasurement,
  TextMeasurement,
} from "./measureTypes";
import { toPaintedText } from "./textMeasurementPolicy";

/**
 * Swappable text-measurement backend. A canvas implementation is installed at
 * the composition roots; a Rust/WASM or headless provider can be installed via
 * `setMeasureProvider` without the layout engine importing canvas directly.
 */
export type MeasureProvider = {
  getFontMetrics: (style: FontStyle) => FontMetrics;
  measureTextWidth: (text: string, style: FontStyle) => number;
  measureText: (text: string, style: FontStyle) => TextMeasurement;
  measureRun: (text: string, style: FontStyle) => RunMeasurement;
};

const throwingMeasureProvider: MeasureProvider = {
  getFontMetrics: () => noProvider(),
  measureTextWidth: () => noProvider(),
  measureText: () => noProvider(),
  measureRun: () => noProvider(),
};

function noProvider(): never {
  panic(
    "No MeasureProvider installed — install canvasMeasureProvider " +
      "(the React editor and the test preload do this).",
  );
}

let activeMeasureProvider: MeasureProvider = throwingMeasureProvider;

export const getMeasureProvider = (): MeasureProvider => activeMeasureProvider;

export const setMeasureProvider = (provider: MeasureProvider): void => {
  activeMeasureProvider = provider;
};

export const resetMeasureProvider = (): void => {
  activeMeasureProvider = throwingMeasureProvider;
};

export const getFontMetrics = (style: FontStyle): FontMetrics =>
  activeMeasureProvider.getFontMetrics(style);

export const measureTextWidth = (text: string, style: FontStyle): number => {
  const paintedText = toPaintedText(text);
  let width = activeMeasureProvider.measureTextWidth(paintedText, style);
  for (const gap of automaticTextSpacingGaps(paintedText, style)) {
    width += gap.advancePx;
  }
  return width;
};

export const measureText = (
  text: string,
  style: FontStyle,
): TextMeasurement => {
  const paintedText = toPaintedText(text);
  const measured = activeMeasureProvider.measureText(paintedText, style);
  const gaps = automaticTextSpacingGaps(paintedText, style);
  if (gaps.length === 0) return measured;
  let width = measured.width;
  for (const gap of gaps) width += gap.advancePx;
  return { ...measured, width };
};

export const measureRun = (text: string, style: FontStyle): RunMeasurement => {
  const paintedText = toPaintedText(text);
  const measured = activeMeasureProvider.measureRun(paintedText, style);
  const gaps = automaticTextSpacingGaps(paintedText, style);
  if (gaps.length === 0) return measured;
  const charWidths = [...measured.charWidths];
  let width = measured.width;
  for (const [offset, advance] of automaticGapCodeUnitAdvances(
    paintedText,
    gaps,
  )) {
    charWidths[offset] = (charWidths[offset] ?? 0) + advance;
    width += advance;
  }
  return { ...measured, charWidths, width };
};

type MeasurePositionedRunOptions = {
  text: string;
  style: FontStyle;
  trailingAutomaticSpacingPx?: number;
};

export const measurePositionedRun = ({
  text,
  style,
  trailingAutomaticSpacingPx = 0,
}: MeasurePositionedRunOptions): RunMeasurement => {
  const measured = measureRun(text, style);
  const gaps = automaticTextSpacingGaps(text, style);
  const gapWidth = gaps.reduce((sum, gap) => sum + gap.advancePx, 0);
  const gapAdvances = automaticGapCodeUnitAdvances(text, gaps);
  const glyphWidth =
    measured.charWidths.reduce((sum, width) => sum + width, 0) - gapWidth;
  const shapedWidth = measureTextWidth(text, style) - gapWidth;
  const scale = glyphWidth > 0 ? shapedWidth / glyphWidth : 1;
  const charWidths = measured.charWidths.map((width, offset) => {
    const gap = gapAdvances.get(offset) ?? 0;
    return (width - gap) * scale + gap;
  });
  if (text.length > 0 && trailingAutomaticSpacingPx > 0) {
    for (const [offset, advance] of automaticGapCodeUnitAdvances(text, [
      { offset: text.length, advancePx: trailingAutomaticSpacingPx },
    ]))
      charWidths[offset] = (charWidths[offset] ?? 0) + advance;
  }
  return {
    ...measured,
    charWidths,
    width: shapedWidth + gapWidth + trailingAutomaticSpacingPx,
  };
};
