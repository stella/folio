/**
 * Diff between two strings, as the segments a redline is drawn from.
 *
 * Tokenises (by default into words and the punctuation marks at their edges,
 * preserving the whitespace as part of each token; see {@link pushRunTokens}),
 * runs an LCS, and returns a left-to-right ordered list of
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
 * 2. A match too short to carry meaning is dropped when changes sit on both
 *    sides of it and at least one of them changes words. At either end of the
 *    string it is the reader's anchor, and
 *    between two punctuation or whitespace edits it is what they were made
 *    around; neither is an island.
 * 3. When what survives is still too fragmented for its length, the whole
 *    paragraph is one replacement ({@link isTooFragmented}). Punctuation and
 *    whitespace edits alone never are.
 *
 * At word granularity the LCS itself prefers words: one matched word outweighs
 * any number of matched marks, and a unique mark is never an anchor.
 *
 * A word token carries the whitespace before it, so a string's first word has
 * none while the same word inside the other string does. Word tokens are
 * therefore matched on their text alone; the rules above judge that
 * alignment, and a whitespace difference around a matched word is marked
 * afterwards as a change of the whitespace only
 * ({@link separateWhitespaceChanges}). Otherwise prefixing "(1) " would strike
 * through the unchanged first word.
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

/** Punctuation and whitespace carry no content except legal section/paragraph signs. */
const SEPARATOR_ONLY = /^[\s\p{P}]*$/u;
const CONTENT_PUNCTUATION = /[§¶]/u;

/**
 * A match of at most this many tokens that carries only punctuation or space is
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

/** Unicode general category P: punctuation marks, dashes, brackets and quotes. */
const PUNCTUATION = /^\p{P}$/u;

/** Keep token storage bounded when a run has an unusually long punctuation edge. */
export const MAX_EDGE_PUNCTUATION_TOKENS = 64;

/** The code point starting at `index`, as a string of one or two UTF-16 units. */
const codePointAt = (value: string, index: number): string =>
  String.fromCodePoint(value.codePointAt(index) ?? 0);

/** The code point ending just before `end`. */
const codePointBefore = (value: string, end: number): string => {
  const low = value.charCodeAt(end - 1);
  const high = value.charCodeAt(end - 2);
  const isPair = low >= 0xdc_00 && low <= 0xdf_ff && high >= 0xd8_00 && high <= 0xdb_ff;
  return isPair ? value.slice(end - 2, end) : value.slice(end - 1, end);
};

/**
 * Split one whitespace-free run into its words and punctuation marks.
 *
 * Punctuation marks (general category P) at either edge of the run are
 * tokens of their own, as Word's compare treats them: `jmění.` becoming `jmění,`
 * changes the mark, not the word. Punctuation inside the run stays in the
 * word, so `d.o.o`, `1.1.2026`, `3.5`, `well-known` and `don't` are single
 * tokens; only their edge marks split off (`d.o.o.` is `d.o.o` + `.`, `b)` is
 * `b` + `)`). A run made only of punctuation is split up to the edge cap.
 * `§` and `¶` are content markers despite their Unicode punctuation category;
 * symbols (general category S, such as `€` or `+`) are not punctuation.
 */
const pushRunTokens = (tokens: string[], value: string, run: TokenRange, prefixStart: number) => {
  let wordStart = run.start;
  let pending = value.slice(prefixStart, run.start);
  let leadingMarks = 0;
  while (wordStart < run.end && leadingMarks < MAX_EDGE_PUNCTUATION_TOKENS) {
    const mark = codePointAt(value, wordStart);
    if (!PUNCTUATION.test(mark) || CONTENT_PUNCTUATION.test(mark)) {
      break;
    }
    tokens.push(pending + mark);
    pending = "";
    wordStart += mark.length;
    leadingMarks++;
  }
  let wordEnd = run.end;
  const trailing = [];
  while (wordEnd > wordStart && trailing.length < MAX_EDGE_PUNCTUATION_TOKENS) {
    const mark = codePointBefore(value, wordEnd);
    if (!PUNCTUATION.test(mark) || CONTENT_PUNCTUATION.test(mark)) {
      break;
    }
    trailing.push(mark);
    wordEnd -= mark.length;
  }
  if (wordStart < wordEnd) {
    tokens.push(pending + value.slice(wordStart, wordEnd));
    pending = "";
  }
  for (let index = trailing.length - 1; index >= 0; index--) {
    tokens.push(pending + (trailing[index] ?? ""));
    pending = "";
  }
};

export const tokenizeWords = (value: string): string[] => {
  const tokens: string[] = [];
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
    const runStart = cursor;
    while (cursor < value.length && !WHITESPACE.test(value.charAt(cursor))) {
      cursor++;
    }
    pushRunTokens(tokens, value, { start: runStart, end: cursor }, tokenStart);
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

const isSeparatorOnly = (text: string): boolean =>
  SEPARATOR_ONLY.test(text) && !CONTENT_PUNCTUATION.test(text);

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

/**
 * How the LCS scores a match. `uniform` counts every matched token once;
 * `content-first` (word granularity) never trades a matched word for matched
 * punctuation, so moving a mark across a word cannot strike the word through.
 */
const MATCH_WEIGHTING = {
  Uniform: "uniform",
  ContentFirst: "content-first",
} as const;

type MatchWeighting = (typeof MATCH_WEIGHTING)[keyof typeof MATCH_WEIGHTING];

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
  weighting: MatchWeighting;
};

type AlignmentResult = {
  runs: DiffRun[];
  monotoneDirection: "insertion" | "deletion" | null;
};

const hasOppositeChange = (
  runs: readonly DiffRun[],
  direction: "insertion" | "deletion",
): boolean => runs.some(({ type }) => type === (direction === "insertion" ? "del" : "ins"));

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
type PatienceAnchorOptions = {
  beforeKeys: readonly string[];
  afterKeys: readonly string[];
  beforeRange: TokenRange;
  afterRange: TokenRange;
  weighting: MatchWeighting;
};

const findPatienceAnchors = ({
  beforeKeys,
  afterKeys,
  beforeRange,
  afterRange,
  weighting,
}: PatienceAnchorOptions): TokenAnchor[] => {
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
    // A unique comma is no lineage evidence; under content-first weighting
    // it must not pin the alignment the weighted LCS would reject.
    const eligible = weighting === MATCH_WEIGHTING.Uniform || !isSeparatorOnly(key);
    if (eligible && beforeCount === 1 && afterOccurrence?.count === 1) {
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
  weighting: MatchWeighting;
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
  weighting,
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
  // One content match outweighs every separator match the gap could hold, so
  // the LCS keeps all the words it can and only then the marks between them.
  // The cell budget bounds the shorter side at 2,000 tokens, so the largest
  // score, about 2,000 * 2,001, fits a Uint32.
  const contentWeight =
    weighting === MATCH_WEIGHTING.ContentFirst ? Math.min(beforeLength, afterLength) + 1 : 1;
  const matchWeights = new Uint32Array(beforeLength);
  for (let beforeOffset = 0; beforeOffset < beforeLength; beforeOffset++) {
    const key = beforeKeys[beforeRange.start + beforeOffset] ?? "";
    matchWeights[beforeOffset] = isSeparatorOnly(key) ? 1 : contentWeight;
  }
  for (let beforeOffset = 0; beforeOffset < beforeLength; beforeOffset++) {
    const matchWeight = matchWeights[beforeOffset] ?? 1;
    for (let afterOffset = 0; afterOffset < afterLength; afterOffset++) {
      const index = beforeOffset * afterLength + afterOffset;
      const beforeKey = beforeKeys[beforeRange.start + beforeOffset];
      const afterKey = afterKeys[afterRange.start + afterOffset];
      if (beforeKey === afterKey) {
        const diagonal =
          beforeOffset > 0 && afterOffset > 0
            ? (lengths[(beforeOffset - 1) * afterLength + afterOffset - 1] ?? 0)
            : 0;
        lengths[index] = diagonal + matchWeight;
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
  weighting,
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
  const anchors = findPatienceAnchors({
    beforeKeys,
    afterKeys,
    beforeRange: middleRangeBefore,
    afterRange: middleRangeAfter,
    weighting,
  });

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
      weighting,
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
      weighting,
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
    weighting,
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

/** True when a run changes words, not only punctuation or whitespace. */
const isContentChange = (run: DiffRun): boolean =>
  run.type !== "equal" && !isSeparatorOnly(run.text);

type ChangeRegionCursor = { from: number; step: 1 | -1 };

/** True when the change region starting at `from`, walking by `step`, changes words. */
const changeCarriesContent = (
  runs: readonly DiffRun[],
  { from, step }: ChangeRegionCursor,
): boolean => {
  for (let index = from; ; index += step) {
    const run = runs[index];
    if (run === undefined || run.type === "equal") {
      return false;
    }
    if (isContentChange(run)) {
      return true;
    }
  }
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
    // A word between two punctuation or whitespace edits is what those edits
    // were made around, not a coincidence inside a rewrite. One changed word
    // beside it is enough to make it one.
    const isIsland =
      runs[index - 1] !== undefined &&
      runs[index + 1] !== undefined &&
      (changeCarriesContent(runs, { from: index - 1, step: -1 }) ||
        changeCarriesContent(runs, { from: index + 1, step: 1 }));
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

const TOKEN_WHITESPACE = /^(\s*)(.*?)(\s*)$/su;

type TokenParts = { leading: string; text: string; trailing: string };

const splitTokenWhitespace = (token: string): TokenParts => {
  const [, leading = "", text = "", trailing = ""] = TOKEN_WHITESPACE.exec(token) ?? [];
  return { leading, text, trailing };
};

const pushChangedText = (runs: DiffRun[], before: string, after: string): void => {
  if (before.length > 0) {
    pushRun(runs, { type: "del", text: before });
  }
  if (after.length > 0) {
    pushRun(runs, { type: "ins", text: after });
  }
};

const pushWhitespace = (runs: DiffRun[], before: string, after: string): void => {
  if (before !== after) {
    pushChangedText(runs, before, after);
    return;
  }
  if (before.length > 0) {
    pushRun(runs, { type: "equal", before, after, units: 0 });
  }
};

/**
 * Mark the whitespace a matched word gained, lost or changed, leaving the word
 * itself equal. Word tokens are aligned on their text alone, so an `equal` run
 * may pair `"Závislá"` with `" Závislá"`; both sides of such a run tokenize
 * to the same number of words. Every piece of both texts is emitted exactly
 * once and in order, so both strings reconstruct whatever the pairing.
 */
const separateWhitespaceChanges = (runs: readonly DiffRun[]): DiffRun[] => {
  const separated: DiffRun[] = [];
  for (const run of runs) {
    if (run.type !== "equal" || run.before === run.after) {
      pushRun(separated, run);
      continue;
    }
    const beforeTokens = tokenizeWords(run.before);
    const afterTokens = tokenizeWords(run.after);
    const tokenCount = Math.max(beforeTokens.length, afterTokens.length);
    for (let index = 0; index < tokenCount; index++) {
      const before = splitTokenWhitespace(beforeTokens[index] ?? "");
      const after = splitTokenWhitespace(afterTokens[index] ?? "");
      pushWhitespace(separated, before.leading, after.leading);
      if (before.text.length > 0 && after.text.length > 0) {
        pushRun(separated, { type: "equal", before: before.text, after: after.text, units: 1 });
      } else {
        pushChangedText(separated, before.text, after.text);
      }
      pushWhitespace(separated, before.trailing, after.trailing);
    }
  }
  return separated;
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

const WORD_CHARACTER = /^[\p{L}\p{N}\p{M}]$/u;
const LINE_BREAK = /^[\n\r\p{Zl}\p{Zp}]$/u;

/**
 * How well a change edge between two code points reads, after
 * diff-match-patch's semantic score: the string's edge (an empty side), then
 * a line break, then the gap after a sentence or clause mark, then any space,
 * then any other mark; inside a word is worst.
 */
const boundaryScore = (previous: string, next: string): number => {
  if (previous.length === 0 || next.length === 0) {
    return 6;
  }
  if (LINE_BREAK.test(previous) || LINE_BREAK.test(next)) {
    return 4;
  }
  const previousIsSpace = WHITESPACE.test(previous);
  const nextIsSpace = WHITESPACE.test(next);
  if (!previousIsSpace && !WORD_CHARACTER.test(previous) && nextIsSpace) {
    return 3;
  }
  if (previousIsSpace || nextIsSpace) {
    return 2;
  }
  return WORD_CHARACTER.test(previous) && WORD_CHARACTER.test(next) ? 0 : 1;
};

/** True when `position` falls between the two halves of a surrogate pair. */
const splitsSurrogatePair = (text: string, position: number): boolean => {
  const low = text.charCodeAt(position);
  const high = text.charCodeAt(position - 1);
  return low >= 0xdc_00 && low <= 0xdf_ff && high >= 0xd8_00 && high <= 0xdb_ff;
};

type SlideWindow = {
  /** The equal text before the change, the change, and the equal text after it. */
  text: string;
  start: number;
  length: number;
  /** How far the change may slide without emptying an equality it must keep. */
  minimumStart: number;
  maximumEnd: number;
  /** The code points just outside the window on the change's side; empty at the string's edge. */
  outsideBefore: string;
  outsideAfter: string;
};

/**
 * Where, among the lossless positions of one change, it reads best. The
 * change may slide by one character whenever the character it gives up
 * equals the one it takes on, which keeps both strings intact. Only the
 * window decides the result, so an insertion and the deletion that undoes it
 * land on the same text; the leftmost of equally good positions wins.
 */
const bestSlideStart = (window: SlideWindow): number => {
  const { text, start, length, minimumStart, maximumEnd } = window;
  const scoreAt = (position: number): number =>
    boundaryScore(
      position === 0 ? window.outsideBefore : codePointBefore(text, position),
      position === text.length ? window.outsideAfter : codePointAt(text, position),
    );
  let leftmost = start;
  while (leftmost > minimumStart && text[leftmost - 1] === text[leftmost + length - 1]) {
    leftmost--;
  }
  let best = start;
  let bestScore = -1;
  for (let candidate = leftmost; candidate + length <= maximumEnd; candidate++) {
    const end = candidate + length;
    if (!splitsSurrogatePair(text, candidate) && !splitsSurrogatePair(text, end)) {
      const startScore = scoreAt(candidate);
      const endScore = scoreAt(end);
      // A string edge must not outweigh cutting a word the change never touched.
      const splitsWord = candidate !== start && (startScore === 0 || endScore === 0);
      if (!splitsWord && startScore + endScore > bestScore) {
        best = candidate;
        bestScore = startScore + endScore;
      }
    }
    if (end === text.length || text[candidate] !== text[end]) {
      break;
    }
  }
  return best;
};

type SegmentType = WordDiffSegment["type"];

/**
 * Whether a change may end up directly beside `neighbour` once the equality
 * between them empties: always beside nothing, an equality or its own kind,
 * and a deletion may precede an insertion; an insertion before a deletion
 * would break deletion-first order.
 */
const mayAbut = (first: SegmentType | undefined, second: SegmentType | undefined): boolean =>
  first === undefined ||
  second === undefined ||
  first === "equal" ||
  second === "equal" ||
  first === second ||
  (first === "del" && second === "ins");

/** The code point nearest `index`, walking by `step`, on the side `type` belongs to. */
const sideCodePoint = (
  segments: readonly WordDiffSegment[],
  { from, step }: ChangeRegionCursor,
  type: SegmentType,
): string => {
  for (let index = from; index >= 0 && index < segments.length; index += step) {
    const segment = segments[index];
    if (segment === undefined || segment.text.length === 0) {
      continue;
    }
    if (segment.type === "equal" || segment.type === type) {
      return step === 1
        ? codePointAt(segment.text, 0)
        : codePointBefore(segment.text, segment.text.length);
    }
  }
  return "";
};

const mergeAdjacentSegments = (segments: readonly WordDiffSegment[]): WordDiffSegment[] => {
  const merged: WordDiffSegment[] = [];
  for (const segment of segments) {
    const last = merged.at(-1);
    if (segment.text.length === 0) {
      continue;
    }
    if (last?.type === segment.type) {
      last.text += segment.text;
      continue;
    }
    merged.push(segment);
  }
  return merged;
};

/**
 * Slide every insertion or deletion that sits between equalities to the
 * position that reads best.
 *
 * An LCS places a change arbitrarily among equally long alignments: an
 * appended sentence can be marked as `". New sentence"` before the old
 * sentence's full stop, not `" New sentence."` after it. The redline then
 * marks a stop the author never touched and leaves the new sentence's own
 * unmarked. Sliding moves only where a change's edges fall, so both strings
 * still reconstruct.
 */
const slideChangesToReadableBoundaries = (
  segments: readonly WordDiffSegment[],
): WordDiffSegment[] => {
  // Empty equalities at both ends give a change at either edge somewhere to
  // hand the text it slides past.
  const slid: WordDiffSegment[] = [
    { type: "equal", text: "" },
    ...segments.map((segment) => ({ ...segment })),
    { type: "equal", text: "" },
  ];
  for (let index = 1; index < slid.length - 1; index++) {
    const change = slid[index];
    const previous = slid[index - 1];
    const next = slid[index + 1];
    if (
      change === undefined ||
      previous === undefined ||
      next === undefined ||
      change.type === "equal" ||
      previous.type !== "equal" ||
      next.type !== "equal"
    ) {
      continue;
    }
    const text = previous.text + change.text + next.text;
    const start = bestSlideStart({
      text,
      start: previous.text.length,
      length: change.text.length,
      minimumStart: mayAbut(slid[index - 2]?.type, change.type) ? 0 : 1,
      maximumEnd: mayAbut(change.type, slid[index + 2]?.type) ? text.length : text.length - 1,
      outsideBefore: sideCodePoint(slid, { from: index - 2, step: -1 }, change.type),
      outsideAfter: sideCodePoint(slid, { from: index + 2, step: 1 }, change.type),
    });
    const end = start + change.text.length;
    previous.text = text.slice(0, start);
    change.text = text.slice(start, end);
    next.text = text.slice(end);
  }
  return mergeAdjacentSegments(slid);
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
  const requestedNormalization = options.normalization ?? {};
  // A word token's only whitespace is at its edges, so matching with
  // whitespace normalized compares its text alone (see the module header).
  const normalization =
    granularity === "word"
      ? { ...requestedNormalization, whitespace: true }
      : requestedNormalization;
  const whitespaceIsSignificant =
    granularity === "word" && requestedNormalization.whitespace !== true;
  const weighting = granularity === "word" ? MATCH_WEIGHTING.ContentFirst : MATCH_WEIGHTING.Uniform;
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
    weighting,
  });
  let surviving = demoteRejectedMatches(aligned.runs);
  const monotoneDirection = aligned.monotoneDirection;
  if (monotoneDirection !== null && hasOppositeChange(surviving, monotoneDirection)) {
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
        weighting,
      });
      surviving = hasOppositeChange(unanchored, monotoneDirection)
        ? alignMonotoneChange(
            before,
            after,
            beforeTokens,
            afterTokens,
            normalization,
            monotoneDirection,
          )
        : unanchored;
    } else {
      surviving = hasOppositeChange(aligned.runs, monotoneDirection)
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
  // Punctuation and whitespace edits alone are not a rewrite, however many
  // words they leave between them.
  if (
    aligned.monotoneDirection === null &&
    surviving.some(isContentChange) &&
    isTooFragmented(surviving, (before.length + after.length) / 2)
  ) {
    return wholeStringReplacement(before, after);
  }
  let marked = whitespaceIsSignificant ? separateWhitespaceChanges(surviving) : surviving;
  if (monotoneDirection !== null && hasOppositeChange(marked, monotoneDirection)) {
    const monotoneRuns = alignMonotoneChange(
      before,
      after,
      beforeTokens,
      afterTokens,
      normalization,
      monotoneDirection,
    );
    marked = whitespaceIsSignificant ? separateWhitespaceChanges(monotoneRuns) : monotoneRuns;
  }
  const segments = toSegments(orderDeletionsFirst(marked));
  // A normalized equal run carries only the before text, so sliding across
  // it would corrupt the after side.
  const equalRunsAreExact =
    requestedNormalization.case !== true && requestedNormalization.whitespace !== true;
  return granularity === "word" && equalRunsAreExact
    ? slideChangesToReadableBoundaries(segments)
    : segments;
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
