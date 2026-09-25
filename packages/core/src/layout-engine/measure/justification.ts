/**
 * Justified-line fitting: the per-paragraph fit strategy, how far a justified
 * line may compress its spaces, and when shrinking beats wrapping.
 */

import type { ParagraphBlock } from "../types";
import { measureTextWidth } from "./measureProvider";
import type { FontStyle } from "./measureTypes";
import { countCompressibleSpaces } from "./textMeasurementPolicy";
import { WIDTH_TOLERANCE } from "./paragraphMeasureShared";
import type { LineState } from "./paragraphMeasureShared";

const JUSTIFY_SHRINK_TOLERANCE_RATIO = 0.016;
// Modern (`compatibilityMode` 15) justified prose admits a word that overflows
// the measure when every regular space on the line can absorb the overflow by
// shrinking to no less than three quarters of its natural advance.
const JUSTIFY_SPACE_CONTRACTION_RATIO = 0.25;
// Shrinking a justified line's spaces is weighed against stretching the
// shorter line the wrap would leave; see `prefersSpaceShrink`.
const JUSTIFY_STRETCH_PER_SHRINK = 1.7;
const JUSTIFY_ALWAYS_SHRINK_STRETCH = 1.5;
const JUSTIFY_LITERAL_TAB_CONTINUATION_SHRINK_TOLERANCE_RATIO = 0.017;
const JUSTIFY_HANGING_TAB_SHRINK_TOLERANCE_RATIO = 0.021;
const ALL_CAPS_RATIO_THRESHOLD = 0.8;

export function uppercaseLetterRatio(text: string): number {
  let letters = 0;
  let uppercase = 0;
  for (const char of text) {
    const lower = char.toLocaleLowerCase();
    const upper = char.toLocaleUpperCase();
    if (lower === upper) {
      continue;
    }
    letters++;
    if (char === upper) {
      uppercase++;
    }
  }
  return letters === 0 ? 0 : uppercase / letters;
}

type JustifyFitStrategy =
  | { type: "rounding" }
  | { type: "space"; ratio: number }
  | { type: "width"; ratio: number };

export type JustificationProfile = {
  hasTabRuns: boolean;
  uppercaseRatio: number;
};

export function resolveJustifyFitStrategy(
  block: ParagraphBlock,
  isFirstLine: boolean,
  profile: JustificationProfile,
): JustifyFitStrategy {
  if (block.attrs?.justificationCompatibility?.type === "legacy") {
    // No space contraction, but the fit test still absorbs sub-pixel
    // accumulation the way an unjustified paragraph does.
    return { type: "rounding" };
  }
  const hasTabStops = (block.attrs?.tabs?.length ?? 0) > 0;
  if (isFirstLine && profile.hasTabRuns && (block.attrs?.indent?.hanging ?? 0) > 0) {
    return { type: "width", ratio: JUSTIFY_SHRINK_TOLERANCE_RATIO };
  }
  if (!isFirstLine && profile.hasTabRuns && !hasTabStops) {
    return { type: "width", ratio: JUSTIFY_LITERAL_TAB_CONTINUATION_SHRINK_TOLERANCE_RATIO };
  }
  if (profile.hasTabRuns && (isFirstLine || hasTabStops)) {
    return {
      type: "width",
      ratio:
        (block.attrs?.indent?.firstLine ?? 0) === 0
          ? JUSTIFY_HANGING_TAB_SHRINK_TOLERANCE_RATIO
          : JUSTIFY_SHRINK_TOLERANCE_RATIO,
    };
  }
  if (profile.uppercaseRatio > ALL_CAPS_RATIO_THRESHOLD) {
    return { type: "width", ratio: JUSTIFY_SHRINK_TOLERANCE_RATIO };
  }
  return { type: "space", ratio: JUSTIFY_SPACE_CONTRACTION_RATIO };
}

export function compressibleSpaceWidth(text: string, style: FontStyle): number {
  const count = countCompressibleSpaces(text);
  return count === 0 ? 0 : count * measureTextWidth(" ", style);
}

/**
 * Content being fitted onto a justified line: the line width with it added,
 * and the width of spaces inside it that stay on the line (an atomic field
 * result). A word's own trailing space hangs instead, so it is not counted.
 */
type JustifyCandidate = { width: number; innerSpaceWidth: number };

export function justifyFitTolerance(
  line: LineState,
  strategy: JustifyFitStrategy,
  candidateSpaceWidth: number,
  candidate?: JustifyCandidate,
): number {
  if (strategy.type === "rounding") {
    return WIDTH_TOLERANCE;
  }
  if (strategy.type === "width") {
    return Math.max(WIDTH_TOLERANCE, line.availableWidth * strategy.ratio);
  }
  if (candidate !== undefined && !prefersSpaceShrink(line, candidate)) {
    return WIDTH_TOLERANCE;
  }
  return Math.max(WIDTH_TOLERANCE, (line.regularSpaceWidth + candidateSpaceWidth) * strategy.ratio);
}

/**
 * A justified line whose next word overflows the measure has two outcomes:
 * keep the word and shrink the line's spaces, or wrap it and stretch the
 * spaces left on the shorter line. Shrinking wins only when the stretch it
 * avoids is large compared with the shrink it costs: the shrink factor may
 * not exceed `1 + (stretch - 1) / JUSTIFY_STRETCH_PER_SHRINK`, and a stretch
 * beyond `JUSTIFY_ALWAYS_SHRINK_STRETCH` always prefers shrinking. The space
 * before the wrapped word hangs at the line end, so it does not stretch; a
 * line left with nothing to stretch defers to the shrink budget alone.
 */
function prefersSpaceShrink(line: LineState, candidate: JustifyCandidate): boolean {
  const overflow = candidate.width - line.availableWidth;
  if (overflow <= WIDTH_TOLERANCE) {
    return true;
  }
  const shrinkableSpaceWidth = line.regularSpaceWidth + candidate.innerSpaceWidth;
  if (shrinkableSpaceWidth <= overflow) {
    return false;
  }
  const hangingSpaceWidth = Math.min(line.trailingWhitespaceWidth, line.regularSpaceWidth);
  const stretchableSpaceWidth = line.regularSpaceWidth - hangingSpaceWidth;
  if (stretchableSpaceWidth <= 0) {
    return true;
  }
  const visibleWidth = line.width - line.trailingWhitespaceWidth;
  const stretch = 1 + Math.max(0, line.availableWidth - visibleWidth) / stretchableSpaceWidth;
  if (stretch > JUSTIFY_ALWAYS_SHRINK_STRETCH) {
    return true;
  }
  const shrink = shrinkableSpaceWidth / (shrinkableSpaceWidth - overflow);
  return 1 + (stretch - 1) / JUSTIFY_STRETCH_PER_SHRINK >= shrink;
}
