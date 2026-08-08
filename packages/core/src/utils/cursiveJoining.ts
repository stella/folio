/**
 * Cursive-joining queries for scripts whose letters change shape by position.
 *
 * In Arabic, Syriac, N'Ko, Adlam and friends a letter's glyph depends on its
 * neighbours: isolated, initial, medial or final. Shaping stops at an element
 * boundary whenever the two sides resolve to a different font face, so a run
 * split inside a word — bold on one letter, a different size, a different
 * family — makes the word come apart visually, while Word joins straight
 * through it. These predicates say whether a given split severed a connection,
 * so the painter and the measurer can repair it identically.
 *
 * A boundary where both sides share a face is NOT a problem: browsers shape
 * across inline boxes when no shaping-relevant property changes, so ordinary
 * colour or underline marks (tracked changes, comment anchors) already join.
 *
 * The tables come from the pinned UCD via `scripts/generate-joining-types.ts`;
 * see {@link joiningTypes.gen} for why they are generated rather than authored.
 */

import {
  JOINING_LETTER_RANGES,
  JOINING_TRANSPARENT_RANGES,
  JOINS_BACKWARD_RANGES,
  JOINS_FORWARD_RANGES,
} from "./joiningTypes.gen";

/**
 * Membership test over a flat sorted `[start, end, …]` inclusive-range list.
 *
 * Binary search rather than `.some()`: the transparent table alone is 375
 * ranges, and these run per code point on every measured word.
 */
function inRanges(ranges: readonly number[], cp: number): boolean {
  let low = 0;
  let high = ranges.length / 2 - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    // SAFETY: mid is within [0, ranges.length/2), so both indices exist.
    const start = ranges[mid * 2]!;
    const end = ranges[mid * 2 + 1]!;
    if (cp < start) {
      high = mid - 1;
    } else if (cp > end) {
      low = mid + 1;
    } else {
      return true;
    }
  }
  return false;
}

/** Joining_Type D, L or C: this code point can connect to the one after it. */
export const joinsToFollowing = (cp: number): boolean => inRanges(JOINS_FORWARD_RANGES, cp);

/** Joining_Type D, R or C: this code point can connect to the one before it. */
export const joinsToPreceding = (cp: number): boolean => inRanges(JOINS_BACKWARD_RANGES, cp);

/**
 * Joining_Type T: invisible to joining. Combining marks (harakat, Syriac
 * vowels) sit between two letters without breaking their connection, so they
 * must be skipped when looking for a joining neighbour.
 */
export const isJoiningTransparent = (cp: number): boolean =>
  inRanges(JOINING_TRANSPARENT_RANGES, cp);

/** Joining_Type D, L or R: a cursive letter, as opposed to ZWJ or punctuation. */
export const isCursiveLetter = (cp: number): boolean => inRanges(JOINING_LETTER_RANGES, cp);

/**
 * Whether the text contains any cursive letter at all.
 *
 * Callers gate on this so the all-Latin hot path never pays for joining
 * analysis, mirroring how `hasCjk` gates script segmentation.
 */
export function hasCursiveLetter(text: string): boolean {
  for (const ch of text) {
    // SAFETY: for...of over a string yields whole code points.
    if (isCursiveLetter(ch.codePointAt(0)!)) {
      return true;
    }
  }
  return false;
}

/** Last code point of `text` that joining can see, skipping transparents. */
function lastJoiningRelevantCodePoint(text: string): number | undefined {
  let index = text.length;
  while (index > 0) {
    const unit = text.charCodeAt(index - 1);
    const isLowSurrogate = unit >= 0xdc_00 && unit <= 0xdf_ff;
    const high = isLowSurrogate ? text.charCodeAt(index - 2) : Number.NaN;
    const isPair = high >= 0xd8_00 && high <= 0xdb_ff;
    const start = isLowSurrogate && isPair ? index - 2 : index - 1;
    // SAFETY: start is a valid code-point boundary within the string.
    const cp = text.codePointAt(start)!;
    if (!isJoiningTransparent(cp)) {
      return cp;
    }
    index = start;
  }
  return undefined;
}

/** First code point of `text` that joining can see, skipping transparents. */
function firstJoiningRelevantCodePoint(text: string): number | undefined {
  for (const ch of text) {
    // SAFETY: for...of over a string yields whole code points.
    const cp = ch.codePointAt(0)!;
    if (!isJoiningTransparent(cp)) {
      return cp;
    }
  }
  return undefined;
}

/**
 * Whether `before` and `after` were cursively connected before something split
 * them apart.
 *
 * True means the two sides need a joiner to keep rendering as one word. False
 * means the split is harmless: a space, a non-joining letter such as alef on
 * the left of the cut, punctuation, or a script that does not join at all.
 */
export function joinsAcrossBoundary(before: string, after: string): boolean {
  const previous = lastJoiningRelevantCodePoint(before);
  if (previous === undefined || !joinsToFollowing(previous)) {
    return false;
  }
  const next = firstJoiningRelevantCodePoint(after);
  return next !== undefined && joinsToPreceding(next);
}
