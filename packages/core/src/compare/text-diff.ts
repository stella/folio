/**
 * Pure bounded text comparison for representation-neutral content.
 *
 * Tokenises (by default on whitespace boundaries, preserving the whitespace as
 * part of each token), runs an LCS, and returns a left-to-right ordered list of
 * segments where shared runs render as `equal`, removed runs as `del`, and
 * added runs as `ins`. Used by the panel (to render minimal-change redlines),
 * the version comparison, and the apply engine (so tracked changes mark only
 * the divergent tokens, not the whole replaced span).
 *
 * ## The shortest diff is not the most readable one
 *
 * An LCS maximises matched characters, which on a rewritten sentence means
 * matching every stray "the" and comma it can reach. The reader then gets a
 * shredded paragraph — a dozen struck-through fragments interleaved with a
 * dozen inserted ones — where one deletion followed by one insertion says the
 * same thing and can actually be read. Three rules pull the output back:
 *
 * 1. A match made only of separators is not a match ({@link isSeparatorOnly}).
 * 2. A match too short to carry meaning is dropped unless it opens the string,
 *    where it is the reader's anchor rather than an island.
 * 3. When what survives is still too fragmented for its length, the whole
 *    paragraph is one replacement ({@link isTooFragmented}).
 *
 * Common affixes and unique-token anchors split the input into independent
 * gaps before any quadratic work. The residual gaps share one
 * {@link MAX_WORD_DIFF_CELLS} allowance per comparison or apply scope; once it
 * is spent, a gap is a single `del` + `ins` pair. A standalone call owns one
 * fresh scope.
 */

export type WordDiffSegment = {
  type: "equal" | "del" | "ins";
  text: string;
};

/**
 * What one token is. `"word"` tokenises on whitespace and is what a redline
 * over prose should use; `"character"` marks the changed letters inside a
 * word, which reads well for a reference number or a date and badly for a
 * sentence.
 */
export const WORD_DIFF_GRANULARITIES = Object.freeze(["word", "character"] as const);

export type WordDiffGranularity = (typeof WORD_DIFF_GRANULARITIES)[number];

/**
 * Differences the caller does not want marked.
 *
 * A normalized run is reported as `equal` and carries the BEFORE string's
 * text, so the before side still reconstructs exactly while the after side
 * reconstructs only up to the normalization. A caller that must reproduce the
 * after string — anything generating tracked changes — leaves both off.
 */
export type WordDiffNormalization = {
  /** `"Shall"` and `"shall"` are the same token. */
  case?: boolean;
  /** Whitespace around and inside a token does not distinguish it. */
  whitespace?: boolean;
};

export type WordDiffOptions = {
  /** Default `"word"`. */
  granularity?: WordDiffGranularity;
  /** Default: nothing normalized, so both strings reconstruct exactly. */
  normalization?: WordDiffNormalization;
};

const WHITESPACE = /\s/u;

/** Punctuation, symbols and whitespace: everything that is not content. */
const SEPARATOR_ONLY = /^[\s\p{P}\p{S}]*$/u;

/**
 * A match of at most this many tokens that carries no letters or digits is
 * noise: a lone space, a comma, a stray closing bracket. Matching it splits
 * two rewrites into four.
 */
const MAX_SEPARATOR_ONLY_MATCH_UNITS = 3;

/**
 * A match shorter than this, with changes on BOTH sides of it, is dropped into
 * them. An island that small is almost always a coincidence, and it costs the
 * reader two extra fragments to notice. A match that opens or closes the
 * string is not an island: it is where the reader anchors, and striking it
 * through to re-insert it identically reads as an edit nobody made.
 */
const MINIMUM_ISOLATED_MATCH_UNITS = 2;

/**
 * Fragmentation floor. `sumOfSquares` rewards few long matches and punishes
 * many short ones — a single run of length L scores L^2, while L runs of
 * length 1 score L — so comparing it against the average string length
 * separates "a few words changed" from "rewritten, with coincidental
 * matches". Below the floor the paragraph is replaced whole.
 */
const FRAGMENTATION_SCALE = 32;

/** Below this length every match is a large fraction of the string, so the floor says nothing. */
const FRAGMENTATION_MINIMUM_AVERAGE_LENGTH = 8;

const tokenizeWords = (value: string): string[] => {
  const tokens = [];
  let tokenStart = 0;
  let cursor = 0;
  // Walk once instead of backtracking across untrusted whitespace-only document text.
  while (cursor < value.length) {
    while (cursor < value.length && WHITESPACE.test(value.charAt(cursor))) {
      cursor++;
    }
    if (cursor === value.length) {
      break;
    }
    while (cursor < value.length && !WHITESPACE.test(value.charAt(cursor))) {
      cursor++;
    }
    tokens.push(value.slice(tokenStart, cursor));
    tokenStart = cursor;
  }

  const last = tokens.at(-1);
  if (last === undefined) {
    return value.length === 0 ? [] : [value];
  }
  if (tokenStart < value.length) {
    tokens[tokens.length - 1] = last + value.slice(tokenStart);
  }
  return tokens;
};

/** Code points, not UTF-16 units, so an emoji or a surrogate pair stays whole. */
const tokenizeCharacters = (value: string): string[] => [...value];

const tokenize = (value: string, granularity: WordDiffGranularity): string[] =>
  granularity === "character" ? tokenizeCharacters(value) : tokenizeWords(value);

/**
 * The key two tokens are matched on. `toLowerCase` rather than
 * `toLocaleLowerCase`: the ambient locale would make the same two documents
 * diff differently on two machines.
 */
const comparisonKey = (token: string, normalization: WordDiffNormalization): string => {
  const collapsed =
    normalization.whitespace === true ? token.replaceAll(/\s+/gu, " ").trim() : token;
  return normalization.case === true ? collapsed.toLowerCase() : collapsed;
};

const isSeparatorOnly = (text: string): boolean => SEPARATOR_ONLY.test(text);

/**
 * Cell budget shared by the residual LCS gaps inside one comparison or apply
 * scope. `before`/`after` come from attacker-controlled document text (a
 * `modified` block pair), so an unbounded pair of large strings would
 * otherwise force a quadratic-sized allocation. Package-level callers share
 * one scope across every story and operation; the standalone public helper
 * gets one scope per call.
 */
const MAX_WORD_DIFF_CELLS = 4_000_000;

/**
 * Maximum combined residual tokens considered for unique-anchor discovery.
 * Comparison-key arrays, occurrence maps, and the candidate list are optional
 * linear storage over attacker-controlled text. Common affixes are removed
 * before this cap, so a small edit in a very long paragraph can still use
 * anchors without duplicating the whole input as normalized strings.
 */
const MAX_WORD_DIFF_ANCHOR_TOKENS = 16_384;

/**
 * Maximum source code units retained in comparison-key arrays. Token count
 * alone is not a storage bound: one word token can itself contain megabytes,
 * and case or whitespace normalization creates another string for it.
 */
const MAX_WORD_DIFF_COMPARISON_KEY_CODE_UNITS = 1_048_576;

const ALIGNMENT_OPERATION = {
  Equal: 1,
  Delete: 2,
  Insert: 3,
} as const;

/** One aligned run, before the quality rules turn matches into changes. */
type DiffRun =
  | { type: "equal"; before: string; after: string; units: number }
  | { type: "del"; text: string }
  | { type: "ins"; text: string };

type AlignOptions = {
  beforeText: string;
  afterText: string;
  before: readonly string[];
  after: readonly string[];
  normalization: WordDiffNormalization;
  budget: AlignmentBudget;
};

type AlignmentResult = {
  runs: DiffRun[];
  monotoneDirection: "insertion" | "deletion" | null;
};

type TokenRange = {
  start: number;
  end: number;
};

type TokenAnchor = {
  beforeIndex: number;
  afterIndex: number;
};

type AlignmentBudget = {
  remainingCells: number;
};

const fitsCellBudget = (
  beforeLength: number,
  afterLength: number,
  budget: AlignmentBudget,
): boolean =>
  beforeLength === 0 ||
  afterLength === 0 ||
  beforeLength <= Math.floor(budget.remainingCells / afterLength);

const pushRun = (runs: DiffRun[], run: DiffRun): void => {
  const last = runs.at(-1);
  if (last?.type === "equal" && run.type === "equal") {
    last.before += run.before;
    last.after += run.after;
    last.units += run.units;
    return;
  }
  if (
    (last?.type === "del" && run.type === "del") ||
    (last?.type === "ins" && run.type === "ins")
  ) {
    last.text += run.text;
    return;
  }
  runs.push(run);
};

const joinTokenRange = (tokens: readonly string[], { start, end }: TokenRange): string => {
  return tokens.slice(start, end).join("");
};

const pushEqualRange = (
  runs: DiffRun[],
  before: readonly string[],
  after: readonly string[],
  beforeRange: TokenRange,
  afterRange: TokenRange,
): void => {
  const units = beforeRange.end - beforeRange.start;
  if (units === 0) {
    return;
  }
  pushRun(runs, {
    type: "equal",
    before: joinTokenRange(before, beforeRange),
    after: joinTokenRange(after, afterRange),
    units,
  });
};

const pushChangedRange = (
  runs: DiffRun[],
  before: readonly string[],
  after: readonly string[],
  beforeRange: TokenRange,
  afterRange: TokenRange,
): void => {
  if (beforeRange.start !== beforeRange.end) {
    pushRun(runs, {
      type: "del",
      text: joinTokenRange(before, beforeRange),
    });
  }
  if (afterRange.start !== afterRange.end) {
    pushRun(runs, {
      type: "ins",
      text: joinTokenRange(after, afterRange),
    });
  }
};

/**
 * Unique tokens common to both ranges, reduced to a monotone subsequence of
 * target positions. Repeated boilerplate is deliberately ineligible: a unique
 * clause number or name is stronger lineage evidence than another occurrence
 * of "the" chosen by an arbitrary LCS tie.
 */
const findPatienceAnchors = (
  beforeKeys: readonly string[],
  afterKeys: readonly string[],
  beforeRange: TokenRange,
  afterRange: TokenRange,
): TokenAnchor[] => {
  const beforeOccurrences = new Map<string, number>();
  for (let index = beforeRange.start; index < beforeRange.end; index++) {
    const key = beforeKeys[index] ?? "";
    beforeOccurrences.set(key, (beforeOccurrences.get(key) ?? 0) + 1);
  }

  const afterOccurrences = new Map<string, { count: number; index: number }>();
  for (let index = afterRange.start; index < afterRange.end; index++) {
    const key = afterKeys[index] ?? "";
    const occurrence = afterOccurrences.get(key);
    if (occurrence) {
      occurrence.count++;
    } else {
      afterOccurrences.set(key, { count: 1, index });
    }
  }

  const candidates: TokenAnchor[] = [];
  for (let beforeIndex = beforeRange.start; beforeIndex < beforeRange.end; beforeIndex++) {
    const key = beforeKeys[beforeIndex] ?? "";
    const beforeCount = beforeOccurrences.get(key);
    const afterOccurrence = afterOccurrences.get(key);
    if (beforeCount === 1 && afterOccurrence?.count === 1) {
      candidates.push({ beforeIndex, afterIndex: afterOccurrence.index });
    }
  }
  if (candidates.length < 2) {
    return candidates;
  }

  const predecessors = new Int32Array(candidates.length);
  predecessors.fill(-1);
  const tailCandidateIndexes = new Int32Array(candidates.length);
  let tailCount = 0;

  for (const [candidateIndex, candidate] of candidates.entries()) {
    let low = 0;
    let high = tailCount;
    while (low < high) {
      const middle = low + Math.floor((high - low) / 2);
      const tail = candidates[tailCandidateIndexes[middle] ?? -1];
      if (tail && tail.afterIndex < candidate.afterIndex) {
        low = middle + 1;
      } else {
        high = middle;
      }
    }
    if (low > 0) {
      predecessors[candidateIndex] = tailCandidateIndexes[low - 1] ?? -1;
    }
    tailCandidateIndexes[low] = candidateIndex;
    if (low === tailCount) {
      tailCount++;
    }
  }

  const anchors: TokenAnchor[] = [];
  let candidateIndex = tailCandidateIndexes[tailCount - 1] ?? -1;
  while (candidateIndex >= 0) {
    const candidate = candidates[candidateIndex];
    if (candidate) {
      anchors.push(candidate);
    }
    candidateIndex = predecessors[candidateIndex] ?? -1;
  }
  anchors.reverse();
  return anchors;
};

type MonotoneChange = "equivalent" | "insertion" | "deletion" | "mixed";

const keysEqual = (before: string, after: string, normalization: WordDiffNormalization): boolean =>
  comparisonKey(before, normalization) === comparisonKey(after, normalization);

type TokenSubsequenceOptions = {
  subsequence: readonly string[];
  sequence: readonly string[];
  subsequenceRange: TokenRange;
  sequenceRange: TokenRange;
  normalization: WordDiffNormalization;
};

const isTokenSubsequence = ({
  subsequence,
  sequence,
  subsequenceRange,
  sequenceRange,
  normalization,
}: TokenSubsequenceOptions): boolean => {
  if (subsequenceRange.start === subsequenceRange.end) {
    return true;
  }
  let subsequenceIndex = subsequenceRange.start;
  let subsequenceKey = comparisonKey(subsequence[subsequenceIndex] ?? "", normalization);
  for (
    let sequenceIndex = sequenceRange.start;
    sequenceIndex < sequenceRange.end;
    sequenceIndex++
  ) {
    if (subsequenceIndex === subsequenceRange.end) {
      return true;
    }
    const token = sequence[sequenceIndex] ?? "";
    if (subsequenceKey !== comparisonKey(token, normalization)) {
      continue;
    }
    subsequenceIndex++;
    if (subsequenceIndex < subsequenceRange.end) {
      const nextToken = subsequence[subsequenceIndex] ?? "";
      subsequenceKey = comparisonKey(nextToken, normalization);
    }
  }
  return subsequenceIndex === subsequenceRange.end;
};

type ClassifyMonotoneChangeOptions = {
  before: readonly string[];
  after: readonly string[];
  beforeRange: TokenRange;
  afterRange: TokenRange;
  normalization: WordDiffNormalization;
};

/** Classify edits whose shorter token sequence survives whole and in order. */
const classifyMonotoneChange = ({
  before,
  after,
  beforeRange,
  afterRange,
  normalization,
}: ClassifyMonotoneChangeOptions): MonotoneChange => {
  const beforeLength = beforeRange.end - beforeRange.start;
  const afterLength = afterRange.end - afterRange.start;
  if (beforeLength === afterLength) {
    for (let offset = 0; offset < beforeLength; offset++) {
      if (
        !keysEqual(
          before[beforeRange.start + offset] ?? "",
          after[afterRange.start + offset] ?? "",
          normalization,
        )
      ) {
        return "mixed";
      }
    }
    return "equivalent";
  }
  if (beforeLength < afterLength) {
    return isTokenSubsequence({
      subsequence: before,
      sequence: after,
      subsequenceRange: beforeRange,
      sequenceRange: afterRange,
      normalization,
    })
      ? "insertion"
      : "mixed";
  }
  return isTokenSubsequence({
    subsequence: after,
    sequence: before,
    subsequenceRange: afterRange,
    sequenceRange: beforeRange,
    normalization,
  })
    ? "deletion"
    : "mixed";
};

const alignMonotoneChange = (
  beforeText: string,
  afterText: string,
  before: readonly string[],
  after: readonly string[],
  normalization: WordDiffNormalization,
  change: Exclude<MonotoneChange, "mixed" | "equivalent">,
): DiffRun[] => {
  const beforeIsSubsequence = change === "insertion";
  const subsequence = beforeIsSubsequence ? before : after;
  const sequence = beforeIsSubsequence ? after : before;
  const sequenceText = beforeIsSubsequence ? afterText : beforeText;
  const runs: DiffRun[] = [];
  let subsequenceIndex = 0;
  let sequenceOffset = 0;
  let sequenceChangeStartOffset = 0;

  for (let sequenceIndex = 0; sequenceIndex < sequence.length; sequenceIndex++) {
    const subsequenceToken = subsequence[subsequenceIndex];
    const sequenceToken = sequence[sequenceIndex];
    const sequenceTokenStart = sequenceOffset;
    sequenceOffset += sequenceToken?.length ?? 0;
    if (
      subsequenceToken === undefined ||
      sequenceToken === undefined ||
      !keysEqual(subsequenceToken, sequenceToken, normalization)
    ) {
      continue;
    }

    if (sequenceChangeStartOffset < sequenceTokenStart) {
      const changedText = sequenceText.slice(sequenceChangeStartOffset, sequenceTokenStart);
      pushRun(
        runs,
        beforeIsSubsequence
          ? { type: "ins", text: changedText }
          : { type: "del", text: changedText },
      );
    }
    pushRun(runs, {
      type: "equal",
      before: beforeIsSubsequence ? subsequenceToken : sequenceToken,
      after: beforeIsSubsequence ? sequenceToken : subsequenceToken,
      units: 1,
    });
    subsequenceIndex++;
    sequenceChangeStartOffset = sequenceOffset;
  }

  if (sequenceChangeStartOffset < sequenceText.length) {
    const changedText = sequenceText.slice(sequenceChangeStartOffset);
    pushRun(
      runs,
      beforeIsSubsequence ? { type: "ins", text: changedText } : { type: "del", text: changedText },
    );
  }
  return runs;
};

type DenseGapOptions = {
  before: readonly string[];
  after: readonly string[];
  beforeKeys: readonly string[];
  afterKeys: readonly string[];
  beforeRange: TokenRange;
  afterRange: TokenRange;
  runs: DiffRun[];
  budget: AlignmentBudget;
};

/** Longest-common-subsequence alignment for one residual, bounded gap. */
const alignDenseGap = ({
  before,
  after,
  beforeKeys,
  afterKeys,
  beforeRange,
  afterRange,
  runs,
  budget,
}: DenseGapOptions): void => {
  const beforeLength = beforeRange.end - beforeRange.start;
  const afterLength = afterRange.end - afterRange.start;
  if (beforeLength === 0 || afterLength === 0) {
    pushChangedRange(runs, before, after, beforeRange, afterRange);
    return;
  }

  if (!fitsCellBudget(beforeLength, afterLength, budget)) {
    pushChangedRange(runs, before, after, beforeRange, afterRange);
    return;
  }
  const cellCount = beforeLength * afterLength;
  budget.remainingCells -= cellCount;

  // No nested JS arrays: at the maximum allowance this table is exactly
  // 16 MB, and the operation trace below is at most m+n bytes.
  const lengths = new Uint32Array(cellCount);
  for (let beforeOffset = 0; beforeOffset < beforeLength; beforeOffset++) {
    for (let afterOffset = 0; afterOffset < afterLength; afterOffset++) {
      const index = beforeOffset * afterLength + afterOffset;
      const beforeKey = beforeKeys[beforeRange.start + beforeOffset];
      const afterKey = afterKeys[afterRange.start + afterOffset];
      if (beforeKey === afterKey) {
        const diagonal =
          beforeOffset > 0 && afterOffset > 0
            ? (lengths[(beforeOffset - 1) * afterLength + afterOffset - 1] ?? 0)
            : 0;
        lengths[index] = diagonal + 1;
        continue;
      }
      const above =
        beforeOffset > 0 ? (lengths[(beforeOffset - 1) * afterLength + afterOffset] ?? 0) : 0;
      const left = afterOffset > 0 ? (lengths[index - 1] ?? 0) : 0;
      lengths[index] = Math.max(above, left);
    }
  }

  const reversedOperations = new Uint8Array(beforeLength + afterLength);
  let operationCount = 0;
  let beforeOffset = beforeLength - 1;
  let afterOffset = afterLength - 1;
  while (beforeOffset >= 0 && afterOffset >= 0) {
    if (
      beforeKeys[beforeRange.start + beforeOffset] === afterKeys[afterRange.start + afterOffset]
    ) {
      reversedOperations[operationCount++] = ALIGNMENT_OPERATION.Equal;
      beforeOffset--;
      afterOffset--;
      continue;
    }
    const above =
      beforeOffset > 0 ? (lengths[(beforeOffset - 1) * afterLength + afterOffset] ?? 0) : 0;
    const left = afterOffset > 0 ? (lengths[beforeOffset * afterLength + afterOffset - 1] ?? 0) : 0;
    // Preserve the old LCS tie break. Because this trace is reversed, choosing
    // insertion on a tie yields deletion before insertion in forward order.
    if (above > left) {
      reversedOperations[operationCount++] = ALIGNMENT_OPERATION.Delete;
      beforeOffset--;
    } else {
      reversedOperations[operationCount++] = ALIGNMENT_OPERATION.Insert;
      afterOffset--;
    }
  }
  while (beforeOffset >= 0) {
    reversedOperations[operationCount++] = ALIGNMENT_OPERATION.Delete;
    beforeOffset--;
  }
  while (afterOffset >= 0) {
    reversedOperations[operationCount++] = ALIGNMENT_OPERATION.Insert;
    afterOffset--;
  }

  let beforeCursor = beforeRange.start;
  let afterCursor = afterRange.start;
  for (let operationIndex = operationCount - 1; operationIndex >= 0; operationIndex--) {
    const operation = reversedOperations[operationIndex];
    if (operation === ALIGNMENT_OPERATION.Equal) {
      pushRun(runs, {
        type: "equal",
        before: before[beforeCursor] ?? "",
        after: after[afterCursor] ?? "",
        units: 1,
      });
      beforeCursor++;
      afterCursor++;
      continue;
    }
    if (operation === ALIGNMENT_OPERATION.Delete) {
      pushRun(runs, { type: "del", text: before[beforeCursor] ?? "" });
      beforeCursor++;
      continue;
    }
    pushRun(runs, { type: "ins", text: after[afterCursor] ?? "" });
    afterCursor++;
  }
};

/** Factor exact boundary matches before spending the residual DP allowance. */
const alignGap = (options: DenseGapOptions): void => {
  const { beforeKeys, afterKeys, runs } = options;
  let beforeStart = options.beforeRange.start;
  let afterStart = options.afterRange.start;
  const beforeEnd = options.beforeRange.end;
  const afterEnd = options.afterRange.end;

  while (
    beforeStart < beforeEnd &&
    afterStart < afterEnd &&
    beforeKeys[beforeStart] === afterKeys[afterStart]
  ) {
    beforeStart++;
    afterStart++;
  }
  pushEqualRange(
    runs,
    options.before,
    options.after,
    { start: options.beforeRange.start, end: beforeStart },
    { start: options.afterRange.start, end: afterStart },
  );

  let beforeMiddleEnd = beforeEnd;
  let afterMiddleEnd = afterEnd;
  while (
    beforeMiddleEnd > beforeStart &&
    afterMiddleEnd > afterStart &&
    beforeKeys[beforeMiddleEnd - 1] === afterKeys[afterMiddleEnd - 1]
  ) {
    beforeMiddleEnd--;
    afterMiddleEnd--;
  }

  alignDenseGap({
    ...options,
    beforeRange: { start: beforeStart, end: beforeMiddleEnd },
    afterRange: { start: afterStart, end: afterMiddleEnd },
  });
  pushEqualRange(
    runs,
    options.before,
    options.after,
    { start: beforeMiddleEnd, end: beforeEnd },
    { start: afterMiddleEnd, end: afterEnd },
  );
};

/**
 * LCS alignment split at stable, unique-token anchors. Patience anchoring keeps
 * repeated boilerplate from winning a tie over a unique legal term, and makes
 * a small edit inside a long paragraph pay for only its changed gap.
 */
const alignTokens = ({
  beforeText,
  afterText,
  before,
  after,
  normalization,
  budget,
}: AlignOptions): AlignmentResult => {
  let commonPrefixLength = 0;
  let beforePrefixLength = 0;
  let afterPrefixLength = 0;
  while (
    commonPrefixLength < before.length &&
    commonPrefixLength < after.length &&
    keysEqual(before[commonPrefixLength] ?? "", after[commonPrefixLength] ?? "", normalization)
  ) {
    beforePrefixLength += before[commonPrefixLength]?.length ?? 0;
    afterPrefixLength += after[commonPrefixLength]?.length ?? 0;
    commonPrefixLength++;
  }

  let beforeMiddleEnd = before.length;
  let afterMiddleEnd = after.length;
  let beforeSuffixLength = 0;
  let afterSuffixLength = 0;
  while (
    beforeMiddleEnd > commonPrefixLength &&
    afterMiddleEnd > commonPrefixLength &&
    keysEqual(before[beforeMiddleEnd - 1] ?? "", after[afterMiddleEnd - 1] ?? "", normalization)
  ) {
    beforeMiddleEnd--;
    afterMiddleEnd--;
    beforeSuffixLength += before[beforeMiddleEnd]?.length ?? 0;
    afterSuffixLength += after[afterMiddleEnd]?.length ?? 0;
  }

  const beforeRange = { start: commonPrefixLength, end: beforeMiddleEnd };
  const afterRange = { start: commonPrefixLength, end: afterMiddleEnd };
  const monotoneChange = classifyMonotoneChange({
    before,
    after,
    beforeRange,
    afterRange,
    normalization,
  });
  if (monotoneChange === "equivalent") {
    return {
      runs: [
        {
          type: "equal",
          before: beforeText,
          after: afterText,
          units: before.length,
        },
      ],
      monotoneDirection: null,
    };
  }
  const monotoneDirection = monotoneChange === "mixed" ? null : monotoneChange;
  const residualTokenCount =
    beforeRange.end - beforeRange.start + afterRange.end - afterRange.start;
  const runs: DiffRun[] = [];
  const fullInputFits = fitsCellBudget(before.length, after.length, budget);
  const fullInputFitsStorage =
    before.length + after.length <= MAX_WORD_DIFF_ANCHOR_TOKENS &&
    beforeText.length + afterText.length <= MAX_WORD_DIFF_COMPARISON_KEY_CODE_UNITS;
  const residualCodeUnitCount =
    beforeText.length -
    beforePrefixLength -
    beforeSuffixLength +
    afterText.length -
    afterPrefixLength -
    afterSuffixLength;

  if (
    residualTokenCount > MAX_WORD_DIFF_ANCHOR_TOKENS ||
    residualCodeUnitCount > MAX_WORD_DIFF_COMPARISON_KEY_CODE_UNITS
  ) {
    if (commonPrefixLength > 0) {
      pushRun(runs, {
        type: "equal",
        before: beforeText.slice(0, beforePrefixLength),
        after: afterText.slice(0, afterPrefixLength),
        units: commonPrefixLength,
      });
    }
    if (beforePrefixLength + beforeSuffixLength < beforeText.length) {
      pushRun(runs, {
        type: "del",
        text: beforeText.slice(beforePrefixLength, beforeText.length - beforeSuffixLength),
      });
    }
    if (afterPrefixLength + afterSuffixLength < afterText.length) {
      pushRun(runs, {
        type: "ins",
        text: afterText.slice(afterPrefixLength, afterText.length - afterSuffixLength),
      });
    }
    if (beforeSuffixLength > 0) {
      pushRun(runs, {
        type: "equal",
        before: beforeText.slice(beforeText.length - beforeSuffixLength),
        after: afterText.slice(afterText.length - afterSuffixLength),
        units: before.length - beforeMiddleEnd,
      });
    }
    return { runs, monotoneDirection };
  }

  // Only the bounded residual gets duplicate token/key arrays. Huge inputs
  // without a useful affix never materialize an attacker-sized second copy.
  const beforeMiddle = before.slice(beforeRange.start, beforeRange.end);
  const afterMiddle = after.slice(afterRange.start, afterRange.end);
  const beforeKeys = beforeMiddle.map((token) => comparisonKey(token, normalization));
  const afterKeys = afterMiddle.map((token) => comparisonKey(token, normalization));
  const middleRangeBefore = { start: 0, end: beforeMiddle.length };
  const middleRangeAfter = { start: 0, end: afterMiddle.length };
  const anchors = findPatienceAnchors(beforeKeys, afterKeys, middleRangeBefore, middleRangeAfter);

  // Affix factoring changes which occurrence wins an LCS tie. If no stronger
  // unique anchor was selected and the original region fits, preserve the
  // historical alignment exactly; factoring is then only a scalability path.
  if (anchors.length === 0 && fullInputFits && fullInputFitsStorage) {
    alignDenseGap({
      before,
      after,
      beforeKeys: before.map((token) => comparisonKey(token, normalization)),
      afterKeys: after.map((token) => comparisonKey(token, normalization)),
      beforeRange: { start: 0, end: before.length },
      afterRange: { start: 0, end: after.length },
      runs,
      budget,
    });
    return { runs, monotoneDirection };
  }

  pushEqualRange(
    runs,
    before,
    after,
    { start: 0, end: commonPrefixLength },
    { start: 0, end: commonPrefixLength },
  );
  // Anchors below index into the bounded residual arrays, not the full input.
  let beforeStart = 0;
  let afterStart = 0;

  for (const anchor of anchors) {
    alignGap({
      before: beforeMiddle,
      after: afterMiddle,
      beforeKeys,
      afterKeys,
      beforeRange: { start: beforeStart, end: anchor.beforeIndex },
      afterRange: { start: afterStart, end: anchor.afterIndex },
      runs,
      budget,
    });
    pushRun(runs, {
      type: "equal",
      before: beforeMiddle[anchor.beforeIndex] ?? "",
      after: afterMiddle[anchor.afterIndex] ?? "",
      units: 1,
    });
    beforeStart = anchor.beforeIndex + 1;
    afterStart = anchor.afterIndex + 1;
  }

  alignGap({
    before: beforeMiddle,
    after: afterMiddle,
    beforeKeys,
    afterKeys,
    beforeRange: { start: beforeStart, end: beforeMiddle.length },
    afterRange: { start: afterStart, end: afterMiddle.length },
    runs,
    budget,
  });
  pushEqualRange(
    runs,
    before,
    after,
    { start: beforeMiddleEnd, end: before.length },
    { start: afterMiddleEnd, end: after.length },
  );
  return { runs, monotoneDirection };
};

/**
 * True when the surviving matches are too short, relative to the strings they
 * sit in, to be read as anything but coincidence.
 */
const isTooFragmented = (runs: readonly DiffRun[], averageLength: number): boolean => {
  if (averageLength < FRAGMENTATION_MINIMUM_AVERAGE_LENGTH) {
    return false;
  }
  let sumOfSquares = 0;
  for (const run of runs) {
    if (run.type === "equal") {
      sumOfSquares += run.before.length * run.before.length;
    }
  }
  return sumOfSquares * FRAGMENTATION_SCALE < averageLength * averageLength;
};

/**
 * Turn a match the quality rules rejected back into the change it interrupts.
 *
 * Only a match with a change beside it can be rejected: with nothing to
 * absorb it, demoting would invent a deletion and an insertion of the same
 * text where the two strings agree.
 */
const demoteRejectedMatches = (runs: readonly DiffRun[]): DiffRun[] => {
  const kept: DiffRun[] = [];
  for (const [index, run] of runs.entries()) {
    if (run.type !== "equal") {
      kept.push(run);
      continue;
    }
    const interruptsAChange = runs[index - 1] !== undefined || runs[index + 1] !== undefined;
    const isIsland = runs[index - 1] !== undefined && runs[index + 1] !== undefined;
    const rejected =
      (interruptsAChange &&
        run.units <= MAX_SEPARATOR_ONLY_MATCH_UNITS &&
        isSeparatorOnly(run.before)) ||
      (isIsland && run.units < MINIMUM_ISOLATED_MATCH_UNITS);
    if (!rejected) {
      kept.push(run);
      continue;
    }
    kept.push({ type: "del", text: run.before }, { type: "ins", text: run.after });
  }
  return kept;
};

const toSegments = (runs: readonly DiffRun[]): WordDiffSegment[] => {
  const segments: WordDiffSegment[] = [];
  const push = (type: WordDiffSegment["type"], text: string): void => {
    if (text.length === 0) {
      return;
    }
    const last = segments.at(-1);
    if (last?.type === type) {
      last.text += text;
      return;
    }
    segments.push({ type, text });
  };
  // A demoted match leaves a `del` next to the `del` before it, so the pass
  // that orders deletions ahead of insertions runs before they are merged.
  for (const run of runs) {
    if (run.type === "equal") {
      push("equal", run.before);
      continue;
    }
    push(run.type, run.text);
  }
  return segments;
};

/**
 * Deletions before insertions within one changed region. Demoting a match
 * emits `del`, `ins`, `del`, `ins`; a reader wants the whole old text struck
 * through and then the whole new text.
 */
const orderDeletionsFirst = (runs: readonly DiffRun[]): DiffRun[] => {
  const ordered: DiffRun[] = [];
  let deletions = "";
  let insertions = "";
  const flush = (): void => {
    if (deletions.length > 0) {
      ordered.push({ type: "del", text: deletions });
    }
    if (insertions.length > 0) {
      ordered.push({ type: "ins", text: insertions });
    }
    deletions = "";
    insertions = "";
  };
  for (const run of runs) {
    if (run.type === "del") {
      deletions += run.text;
      continue;
    }
    if (run.type === "ins") {
      insertions += run.text;
      continue;
    }
    flush();
    ordered.push(run);
  }
  flush();
  return ordered;
};

const wholeStringReplacement = (before: string, after: string): WordDiffSegment[] => {
  const segments: WordDiffSegment[] = [];
  if (before.length > 0) {
    segments.push({ type: "del", text: before });
  }
  if (after.length > 0) {
    segments.push({ type: "ins", text: after });
  }
  return segments;
};

const diffWordSegmentsWithBudget = (
  before: string,
  after: string,
  options: WordDiffOptions,
  budget: AlignmentBudget,
): WordDiffSegment[] => {
  if (before === after) {
    return before.length === 0 ? [] : [{ type: "equal", text: before }];
  }
  const granularity = options.granularity ?? "word";
  const normalization = options.normalization ?? {};
  const beforeTokens = tokenize(before, granularity);
  const afterTokens = tokenize(after, granularity);
  if (beforeTokens.length === 0 && afterTokens.length === 0) {
    return [];
  }
  if (beforeTokens.length === 0 || afterTokens.length === 0) {
    return wholeStringReplacement(before, after);
  }

  const aligned = alignTokens({
    beforeText: before,
    afterText: after,
    before: beforeTokens,
    after: afterTokens,
    normalization,
    budget,
  });
  let surviving = demoteRejectedMatches(aligned.runs);
  const inventedOppositeChange =
    (aligned.monotoneDirection === "insertion" && surviving.some(({ type }) => type === "del")) ||
    (aligned.monotoneDirection === "deletion" && surviving.some(({ type }) => type === "ins"));
  const monotoneDirection = aligned.monotoneDirection;
  if (inventedOppositeChange && monotoneDirection !== null) {
    const unanchoredFitsStorage =
      beforeTokens.length + afterTokens.length <= MAX_WORD_DIFF_ANCHOR_TOKENS &&
      before.length + after.length <= MAX_WORD_DIFF_COMPARISON_KEY_CODE_UNITS;
    if (unanchoredFitsStorage && fitsCellBudget(beforeTokens.length, afterTokens.length, budget)) {
      const unanchored: DiffRun[] = [];
      alignDenseGap({
        before: beforeTokens,
        after: afterTokens,
        beforeKeys: beforeTokens.map((token) => comparisonKey(token, normalization)),
        afterKeys: afterTokens.map((token) => comparisonKey(token, normalization)),
        beforeRange: { start: 0, end: beforeTokens.length },
        afterRange: { start: 0, end: afterTokens.length },
        runs: unanchored,
        budget,
      });
      surviving = unanchored;
    } else {
      const rawInventsOpposite =
        (monotoneDirection === "insertion" && aligned.runs.some(({ type }) => type === "del")) ||
        (monotoneDirection === "deletion" && aligned.runs.some(({ type }) => type === "ins"));
      surviving = rawInventsOpposite
        ? alignMonotoneChange(
            before,
            after,
            beforeTokens,
            afterTokens,
            normalization,
            monotoneDirection,
          )
        : aligned.runs;
    }
  }
  if (
    aligned.monotoneDirection === null &&
    isTooFragmented(surviving, (before.length + after.length) / 2)
  ) {
    return wholeStringReplacement(before, after);
  }
  return toSegments(orderDeletionsFirst(surviving));
};

/**
 * One internal comparison/apply scope. Every diff shares the same quadratic
 * allowance; public standalone calls below still receive a fresh allowance.
 * Not re-exported from a package entry point.
 */
export const createWordDiffSession = (options: WordDiffOptions = {}) => {
  const resolvedOptions = {
    granularity: options.granularity ?? "word",
    normalization: { ...options.normalization },
  } satisfies WordDiffOptions;
  const budget = { remainingCells: MAX_WORD_DIFF_CELLS };
  return {
    diff: (before: string, after: string): WordDiffSegment[] =>
      diffWordSegmentsWithBudget(before, after, resolvedOptions, budget),
  };
};

const WORD_DIFF_SESSION = Symbol("folio.wordDiffSession");

type ScopedWordDiffOptions = WordDiffOptions & {
  readonly [WORD_DIFF_SESSION]: ReturnType<typeof createWordDiffSession>;
};

const hasWordDiffSession = (options: WordDiffOptions): options is ScopedWordDiffOptions =>
  WORD_DIFF_SESSION in options;

/** Attach one internal work scope without widening a caller-facing option type. */
export const createScopedWordDiffOptions = <Options extends WordDiffOptions>(
  options: Options,
): Options => ({
  ...options,
  [WORD_DIFF_SESSION]: createWordDiffSession(options),
});

/** Reuse an internal scope when present; ordinary public callers get a fresh one. */
export const wordDiffSessionFromOptions = (
  options: WordDiffOptions | undefined,
): ReturnType<typeof createWordDiffSession> => {
  if (options && hasWordDiffSession(options)) {
    return options[WORD_DIFF_SESSION];
  }
  return createWordDiffSession(options);
};

export const diffWordSegments = (
  before: string,
  after: string,
  options: WordDiffOptions = {},
): WordDiffSegment[] =>
  diffWordSegmentsWithBudget(before, after, options, {
    remainingCells: MAX_WORD_DIFF_CELLS,
  });
