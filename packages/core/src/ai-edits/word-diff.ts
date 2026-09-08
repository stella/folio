/**
 * Diff between two strings, as the segments a redline is drawn from.
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
 * {@link MAX_WORD_DIFF_CELLS} allowance per call; once it is spent, a gap is a
 * single `del` + `ins` pair.
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
 * Cell budget shared by the residual LCS gaps inside one
 * {@link diffWordSegments} call. `before`/`after` come from attacker-controlled
 * document text (a `modified` block pair), so an unbounded pair of large
 * strings would otherwise force a quadratic-sized allocation. This is not a
 * document-wide budget: callers diffing several blocks receive one allowance
 * per call.
 */
const MAX_WORD_DIFF_CELLS = 4_000_000;

/**
 * Maximum combined residual tokens considered for unique-anchor discovery.
 * The token and comparison-key arrays are already required by the public
 * operation, but the occurrence maps and candidate list are optional linear
 * storage over attacker-controlled text. Common affixes are removed before
 * this cap, so a small edit in a very long paragraph can still use anchors.
 */
const MAX_WORD_DIFF_ANCHOR_TOKENS = 16_384;

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
  before: readonly string[];
  after: readonly string[];
  normalization: WordDiffNormalization;
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
    before: before.slice(beforeRange.start, beforeRange.end).join(""),
    after: after.slice(afterRange.start, afterRange.end).join(""),
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
      text: before.slice(beforeRange.start, beforeRange.end).join(""),
    });
  }
  if (afterRange.start !== afterRange.end) {
    pushRun(runs, {
      type: "ins",
      text: after.slice(afterRange.start, afterRange.end).join(""),
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

  if (beforeLength > Math.floor(budget.remainingCells / afterLength)) {
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
const alignTokens = ({ before, after, normalization }: AlignOptions): DiffRun[] => {
  const beforeKeys = before.map((token) => comparisonKey(token, normalization));
  const afterKeys = after.map((token) => comparisonKey(token, normalization));
  let commonPrefixLength = 0;
  while (
    commonPrefixLength < before.length &&
    commonPrefixLength < after.length &&
    beforeKeys[commonPrefixLength] === afterKeys[commonPrefixLength]
  ) {
    commonPrefixLength++;
  }

  let beforeMiddleEnd = before.length;
  let afterMiddleEnd = after.length;
  while (
    beforeMiddleEnd > commonPrefixLength &&
    afterMiddleEnd > commonPrefixLength &&
    beforeKeys[beforeMiddleEnd - 1] === afterKeys[afterMiddleEnd - 1]
  ) {
    beforeMiddleEnd--;
    afterMiddleEnd--;
  }

  const beforeRange = { start: commonPrefixLength, end: beforeMiddleEnd };
  const afterRange = { start: commonPrefixLength, end: afterMiddleEnd };
  const residualTokenCount =
    beforeRange.end - beforeRange.start + afterRange.end - afterRange.start;
  const anchors =
    residualTokenCount <= MAX_WORD_DIFF_ANCHOR_TOKENS
      ? findPatienceAnchors(beforeKeys, afterKeys, beforeRange, afterRange)
      : [];
  const budget = { remainingCells: MAX_WORD_DIFF_CELLS };
  const runs: DiffRun[] = [];

  // Affix factoring changes which occurrence wins an LCS tie. If no stronger
  // unique anchor was selected and the original region fits, preserve the
  // historical alignment exactly; factoring is then only a scalability path.
  if (
    anchors.length === 0 &&
    (after.length === 0 || before.length <= Math.floor(budget.remainingCells / after.length))
  ) {
    alignDenseGap({
      before,
      after,
      beforeKeys,
      afterKeys,
      beforeRange: { start: 0, end: before.length },
      afterRange: { start: 0, end: after.length },
      runs,
      budget,
    });
    return runs;
  }

  pushEqualRange(
    runs,
    before,
    after,
    { start: 0, end: commonPrefixLength },
    { start: 0, end: commonPrefixLength },
  );
  let beforeStart = commonPrefixLength;
  let afterStart = commonPrefixLength;

  for (const anchor of anchors) {
    alignGap({
      before,
      after,
      beforeKeys,
      afterKeys,
      beforeRange: { start: beforeStart, end: anchor.beforeIndex },
      afterRange: { start: afterStart, end: anchor.afterIndex },
      runs,
      budget,
    });
    pushRun(runs, {
      type: "equal",
      before: before[anchor.beforeIndex] ?? "",
      after: after[anchor.afterIndex] ?? "",
      units: 1,
    });
    beforeStart = anchor.beforeIndex + 1;
    afterStart = anchor.afterIndex + 1;
  }

  alignGap({
    before,
    after,
    beforeKeys,
    afterKeys,
    beforeRange: { start: beforeStart, end: beforeMiddleEnd },
    afterRange: { start: afterStart, end: afterMiddleEnd },
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
  return runs;
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

export const diffWordSegments = (
  before: string,
  after: string,
  options: WordDiffOptions = {},
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

  const aligned = alignTokens({ before: beforeTokens, after: afterTokens, normalization });
  const surviving = demoteRejectedMatches(aligned);
  if (isTooFragmented(surviving, (before.length + after.length) / 2)) {
    return wholeStringReplacement(before, after);
  }
  return toSegments(orderDeletionsFirst(surviving));
};
