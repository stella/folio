/**
 * Line-fitting helpers: character-level fit search, the hyphenation zone,
 * final-candidate admission, and the trailing whitespace, letter-spacing and
 * hanging-punctuation widths a line break may exclude.
 */

import { getHorizontalScaleFactor } from "../../utils/horizontalScale";
import type { ParagraphBlock, MeasuredLine, Run } from "../types";
import { twipsToPx } from "./measureHelpers";
import { measureTextWidth } from "./measureProvider";
import type { FontStyle } from "./measureTypes";
import { findGraphemeBreaks, isHangingPunctuation } from "./lineBreaks";
import { defaultLineBreakProvider, getLineBreakProvider } from "./lineBreakProvider";
import type { LineBreakPolicy } from "./lineBreakProvider";
import { WIDTH_TOLERANCE, runToFontStyle, isTextRun } from "./paragraphMeasureShared";
import type { LineState } from "./paragraphMeasureShared";

/**
 * Find the longest prefix of `text` that fits within `maxWidth` pixels.
 * Returns the number of characters that fit (at least 1 if `forceMin` is true).
 */
export function findMaxFittingLength(
  text: string,
  style: FontStyle,
  maxWidth: number,
  forceMin: boolean = false,
  policy?: LineBreakPolicy,
): number {
  const boundaries = findGraphemeBreaks(text, policy);
  let lo = 0;
  let hi = boundaries.length - 1;
  let best = 0;
  while (lo <= hi) {
    const mid = Math.floor((lo + hi) / 2);
    // SAFETY: lo/hi are bounded by boundaries.length.
    const boundary = boundaries[mid]!;
    if (measureTextWidth(text.slice(0, boundary), style) <= maxWidth) {
      best = boundary;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return forceMin && best === 0 ? (boundaries.at(0) ?? 0) : best;
}

export function exceedsHyphenationZone(line: LineState, zoneTwips: number): boolean {
  if (line.width <= 0) {
    return true;
  }

  const visibleWidth = Math.max(0, line.width - line.trailingWhitespaceWidth);
  const unusedWidth = Math.max(0, line.availableWidth - visibleWidth);
  return unusedWidth > twipsToPx(zoneTwips) + WIDTH_TOLERANCE;
}

export const isIgnorableFinalTailRun = (run: Run): boolean =>
  run.kind === "renderedPageBreak" || (isTextRun(run) && run.text.length === 0);

export function isFinalTextCandidate(
  block: ParagraphBlock,
  runIndex: number,
  nextBreak: number,
): boolean {
  const currentRun = block.runs[runIndex];
  if (!currentRun || !isTextRun(currentRun) || nextBreak !== currentRun.text.length) {
    return false;
  }
  for (let index = runIndex + 1; index < block.runs.length; index += 1) {
    const run = block.runs[index];
    if (run && isIgnorableFinalTailRun(run)) {
      continue;
    }
    return false;
  }
  return true;
}

type TextCandidateFit =
  | { type: "ordinary"; tolerancePx: number }
  | {
      type: "final-contraction-admitted";
      tolerancePx: number;
      paint: NonNullable<MeasuredLine["justificationPaint"]>;
    };

/**
 * Fit of a text candidate on a justified line: the paragraph's final word may
 * be admitted by shrinking spaces, which then needs a paint plan.
 */
export function resolveTextCandidateFit(
  isFinalCandidate: boolean,
  candidateWidth: number,
  availableWidth: number,
  tolerancePx: number,
): TextCandidateFit {
  if (!isFinalCandidate) {
    return { type: "ordinary", tolerancePx };
  }
  return admitFinalCandidate(candidateWidth, availableWidth, tolerancePx);
}

/**
 * A paragraph's last line is not stretched, so a final word admitted by space
 * contraction needs an explicit paint plan; otherwise it would paint past the
 * measure. Sub-pixel rounding overflow stays on the ordinary path.
 */
function admitFinalCandidate(
  candidateWidth: number,
  availableWidth: number,
  tolerancePx: number,
): TextCandidateFit {
  const contractionPx = candidateWidth - availableWidth;
  if (contractionPx <= WIDTH_TOLERANCE || contractionPx > tolerancePx) {
    return { type: "ordinary", tolerancePx };
  }
  return {
    type: "final-contraction-admitted",
    tolerancePx,
    paint: { type: "space-contraction", contractionPx },
  };
}

export function trimTrailingSpacesAndTabs(text: string): string {
  let end = text.length;
  while (end > 0) {
    const char = text[end - 1];
    if (char !== " " && char !== "\t") {
      break;
    }
    end--;
  }
  return text.slice(0, end);
}

export function measureWordWithTrailingWhitespace(
  word: string,
  style: FontStyle,
): { measuredWord: string; wordWidth: number; fullWordWidth: number } {
  const measuredWord = trimTrailingSpacesAndTabs(word);
  const wordWidth = measureTextWidth(measuredWord, style);
  const trailingWhitespace = word.slice(measuredWord.length);
  if (trailingWhitespace.length === 0) {
    return { measuredWord, wordWidth, fullWordWidth: wordWidth };
  }

  // Tabs have contextual advances, while kerning can change the advance at
  // the split boundary. Keep the exact full-string measurement for both.
  if (style.kerning || trailingWhitespace.includes("\t")) {
    return { measuredWord, wordWidth, fullWordWidth: measureTextWidth(word, style) };
  }

  const trailingWhitespaceWidth = measureTextWidth(trailingWhitespace, style);
  const boundarySpacing =
    measuredWord.length > 0
      ? (style.letterSpacing ?? 0) * getHorizontalScaleFactor(style.horizontalScale)
      : 0;
  return {
    measuredWord,
    wordWidth,
    fullWordWidth: wordWidth + trailingWhitespaceWidth + boundarySpacing,
  };
}

const scaledLetterSpacing = (style: FontStyle): number =>
  (style.letterSpacing ?? 0) * getHorizontalScaleFactor(style.horizontalScale);

/**
 * Character spacing still owed after the last character placed on the line,
 * or 0 when the line has no preceding character at `charIndex` of `runIndex`.
 * Within a run it is the run's own spacing; at a boundary between two text
 * runs it is the preceding run's spacing on its final character.
 */
export function precedingLetterSpacing({
  block,
  line,
  runIndex,
  charIndex,
  runStyle,
}: {
  block: ParagraphBlock;
  line: LineState;
  runIndex: number;
  charIndex: number;
  runStyle: FontStyle;
}): number {
  if (line.fromRun === line.toRun && line.fromChar === line.toChar) {
    return 0;
  }
  if (line.toRun === runIndex && line.toChar === charIndex) {
    return scaledLetterSpacing(runStyle);
  }
  if (charIndex !== 0 || line.toRun !== runIndex - 1) {
    return 0;
  }
  const previousRun = block.runs[runIndex - 1];
  if (
    !previousRun ||
    !isTextRun(previousRun) ||
    previousRun.text.length === 0 ||
    line.toChar !== previousRun.text.length
  ) {
    return 0;
  }
  return scaledLetterSpacing(runToFontStyle(previousRun));
}

function trailingCodePoint(text: string): string | undefined {
  if (text.length === 0) {
    return undefined;
  }
  const trailing = text.charCodeAt(text.length - 1);
  const preceding = text.charCodeAt(text.length - 2);
  const hasSurrogatePair =
    trailing >= 0xdc00 && trailing <= 0xdfff && preceding >= 0xd800 && preceding <= 0xdbff;
  const start = hasSurrogatePair ? text.length - 2 : text.length - 1;
  return text.slice(start);
}

const usesDefaultHangingPunctuationClassifier = (): boolean => {
  const provider = getLineBreakProvider();
  const classifier = provider.isHangingPunctuation ?? defaultLineBreakProvider.isHangingPunctuation;
  return classifier === defaultLineBreakProvider.isHangingPunctuation;
};

export function trailingHangingPunctuationWidth(
  text: string,
  style: FontStyle,
  policy: LineBreakPolicy,
  enabled: boolean,
): number {
  if (!enabled || text.length === 0) {
    return 0;
  }
  const trailing = trailingCodePoint(text);
  if (usesDefaultHangingPunctuationClassifier() && !isHangingPunctuation(trailing ?? "", policy)) {
    return 0;
  }
  const boundaries = findGraphemeBreaks(text, policy);
  const lastStart = boundaries.at(-2) ?? 0;
  const lastGrapheme = text.slice(lastStart);
  if (!isHangingPunctuation(lastGrapheme, policy)) {
    return 0;
  }
  return measureTextWidth(lastGrapheme, style);
}
