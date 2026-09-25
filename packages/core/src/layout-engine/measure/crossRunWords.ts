/**
 * Words that span runs: glue widths between differently formatted runs and
 * the segments of a word collected across several runs.
 */

import { splitTrailingToken } from "../../utils/trailingText";
import type { ParagraphBlock, TextRun } from "../types";
import { resolveEffectiveLineBreakPolicy } from "./effectiveLineBreakPolicy";
import { measureTextWidth } from "./measureProvider";
import type { FontStyle } from "./measureTypes";
import { findHyphenationBreaks, findWordBreaks } from "./lineBreaks";
import { MAX_HYPHENATION_WORD_LENGTH } from "./lineBreakProvider";
import { runToFontStyle, isTextRun, isSpaceOrTab } from "./paragraphMeasureShared";
import { trimTrailingSpacesAndTabs } from "./lineFitting";

/**
 * Width of the unbreakable text glued to the end of each run.
 * A run boundary is not itself a wrap opportunity: adjacent note markers and
 * format-only word splits must wrap as one cluster.
 */
export function computeTrailingGlueWidths(block: ParagraphBlock): number[] {
  const runs = block.runs;
  const widths = Array.from({ length: runs.length }, () => 0);
  for (let index = runs.length - 1; index >= 0; index--) {
    const nextRun = runs[index + 1];
    if (!nextRun || !isTextRun(nextRun)) {
      continue;
    }

    const text = nextRun.text;
    if (!text) {
      widths[index] = widths[index + 1] ?? 0;
      continue;
    }
    if (isSpaceOrTab(text.at(0))) {
      continue;
    }

    const style = runToFontStyle(nextRun);
    const breaks = findWordBreaks(
      text,
      resolveEffectiveLineBreakPolicy({ attrs: block.attrs, run: nextRun }).provider,
    );
    const firstBreak = breaks.at(0);
    const leading =
      firstBreak === undefined ? text : trimTrailingSpacesAndTabs(text.slice(0, firstBreak));
    widths[index] =
      measureTextWidth(leading, style) + (firstBreak === undefined ? (widths[index + 1] ?? 0) : 0);
  }
  return widths;
}

/**
 * First character of the text that follows each run, looking through empty
 * text runs; undefined before a non-text run or at the paragraph end.
 */
export function computeFollowingTextLeads(block: ParagraphBlock): (string | undefined)[] {
  const leads: (string | undefined)[] = Array.from({ length: block.runs.length }, () => undefined);
  let lead: string | undefined;
  for (let index = block.runs.length - 1; index >= 0; index--) {
    leads[index] = lead;
    const run = block.runs[index];
    if (!run || !isTextRun(run)) {
      lead = undefined;
    } else if (run.text) {
      lead = run.text[0];
    }
  }
  return leads;
}

export function computeProtectedCrossRunGlueWidths(block: ParagraphBlock): number[] {
  const widths = Array.from({ length: block.runs.length }, () => 0);
  for (let index = 0; index < block.runs.length; index++) {
    const run = block.runs[index];
    if (!run || !isTextRun(run)) {
      continue;
    }
    const policy = resolveEffectiveLineBreakPolicy({ attrs: block.attrs, run }).provider;
    if (!policy.locale?.toLocaleLowerCase().startsWith("cs")) {
      continue;
    }
    const trailing = splitTrailingToken(run.text ?? "");
    if (!trailing || trailing.token.length !== 1) {
      continue;
    }
    const { token } = trailing;

    let separator = trailing.separator;
    let glueWidth = separator.length > 0 ? measureTextWidth(separator, runToFontStyle(run)) : 0;
    let followingWord = "";
    let unseparatedFollowingText = false;
    for (let nextIndex = index + 1; nextIndex < block.runs.length; nextIndex++) {
      const nextRun = block.runs[nextIndex];
      if (!nextRun || !isTextRun(nextRun)) {
        break;
      }
      const text = nextRun.text ?? "";
      let consumed = "";
      for (const char of text) {
        const isWhitespace = /\s/u.test(char);
        if (followingWord.length > 0 && isWhitespace) {
          break;
        }
        if (separator.length === 0 && !isWhitespace) {
          unseparatedFollowingText = true;
          break;
        }
        consumed += char;
        if (isWhitespace) {
          separator += char;
        } else {
          followingWord += char;
        }
      }
      if (consumed.length > 0) {
        glueWidth += measureTextWidth(consumed, runToFontStyle(nextRun));
      }
      if (unseparatedFollowingText) {
        break;
      }
      if (followingWord.length > 0 && consumed.length < text.length) {
        break;
      }
    }
    if (unseparatedFollowingText || separator.length === 0 || followingWord.length === 0) {
      continue;
    }

    const boundary = token.length + separator.length;
    const probe = token + separator + followingWord;
    const breaks = findWordBreaks(probe, policy);
    if (!breaks.includes(boundary)) {
      widths[index] = glueWidth;
    }
  }
  return widths;
}

type CrossRunWordSegment = {
  runIndex: number;
  startChar: number;
  text: string;
  run: TextRun;
  style: FontStyle;
};

type CrossRunWord = {
  text: string;
  width: number;
  breaks: number[];
  segments: CrossRunWordSegment[];
};

type CollectCrossRunWordOptions = {
  block: ParagraphBlock;
  startRunIndex: number;
  startChar: number;
};

/**
 * Cap on the number of runs `collectCrossRunWord` traverses per call. A
 * paragraph split into many tiny (including empty) runs would otherwise let
 * the search walk the full run list for every word-start position it is
 * invoked from, without the per-word text-length budget ever kicking in.
 */
const MAX_CROSS_RUN_SEGMENTS = 64;

/** Collect one lexical word continued through adjacent formatting runs. */
export function collectCrossRunWord({
  block,
  startRunIndex,
  startChar,
}: CollectCrossRunWordOptions): CrossRunWord | undefined {
  const segments: CrossRunWordSegment[] = [];
  let text = "";
  let width = 0;

  for (let runIndex = startRunIndex; runIndex < block.runs.length; runIndex++) {
    if (runIndex - startRunIndex >= MAX_CROSS_RUN_SEGMENTS) {
      return undefined;
    }
    const run = block.runs[runIndex];
    if (!run || !isTextRun(run)) {
      break;
    }

    const start = runIndex === startRunIndex ? startChar : 0;
    const remainder = run.text.slice(start);
    if (remainder.length === 0) {
      continue;
    }
    if (runIndex !== startRunIndex && isSpaceOrTab(remainder[0])) {
      break;
    }

    const remainingBudget = MAX_HYPHENATION_WORD_LENGTH - text.length;
    const boundedRemainder = remainder.slice(0, remainingBudget + 1);
    const firstBreak = findWordBreaks(
      boundedRemainder,
      resolveEffectiveLineBreakPolicy({ attrs: block.attrs, run }).provider,
    ).at(0);
    const segmentText = trimTrailingSpacesAndTabs(
      firstBreak === undefined ? boundedRemainder : boundedRemainder.slice(0, firstBreak),
    );
    if (segmentText.length === 0) {
      break;
    }
    if (segmentText.length > remainingBudget) {
      return undefined;
    }

    const style = runToFontStyle(run);
    segments.push({ runIndex, startChar: start, text: segmentText, run, style });
    text += segmentText;
    width += measureTextWidth(segmentText, style);

    if (firstBreak !== undefined) {
      break;
    }
  }

  if (segments.length < 2) {
    return undefined;
  }
  const firstRun = segments[0]?.run;
  if (!firstRun) {
    return undefined;
  }
  return {
    text,
    width,
    breaks: findHyphenationBreaks(
      text,
      resolveEffectiveLineBreakPolicy({ attrs: block.attrs, run: firstRun }).provider,
    ),
    segments,
  };
}

export type CrossRunPrefix = {
  width: number;
  endRunIndex: number;
  endChar: number;
  hyphenStyle: FontStyle;
  segments: CrossRunWordSegment[];
};

export function measureCrossRunPrefix(
  word: CrossRunWord,
  breakOffset: number,
): CrossRunPrefix | undefined {
  let width = 0;
  let remaining = breakOffset;
  const measuredSegments: CrossRunWordSegment[] = [];
  for (const segment of word.segments) {
    const length = Math.min(remaining, segment.text.length);
    if (length > 0) {
      width += measureTextWidth(segment.text.slice(0, length), segment.style);
      measuredSegments.push(segment);
    }
    if (remaining <= segment.text.length) {
      return {
        width,
        endRunIndex: segment.runIndex,
        endChar: segment.startChar + remaining,
        hyphenStyle: segment.style,
        segments: measuredSegments,
      };
    }
    remaining -= segment.text.length;
  }
  return undefined;
}
