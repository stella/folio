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
 * O(n*m) on token counts; past {@link MAX_WORD_DIFF_CELLS} the DP is skipped
 * for a single whole-string `del` + `ins` pair.
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
 * Cell budget for the O(n*m) DP table below. `before`/`after` come from
 * attacker-controlled document text (a `modified` block pair), so an
 * unbounded pair of large strings would otherwise force a quadratic-sized
 * allocation. Past this budget, skip the DP and fall back to a single
 * whole-string `del` + `ins` pair — a coarser diff, but O(1) memory.
 */
const MAX_WORD_DIFF_CELLS = 4_000_000;

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

/** The LCS alignment, as runs. Longest common subsequence on comparison keys. */
const alignTokens = ({ before, after, normalization }: AlignOptions): DiffRun[] => {
  const beforeKeys = before.map((token) => comparisonKey(token, normalization));
  const afterKeys = after.map((token) => comparisonKey(token, normalization));
  const m = before.length;
  const n = after.length;

  const dp: number[][] = Array.from({ length: m + 1 }, () =>
    Array.from({ length: n + 1 }, () => 0),
  );
  for (let i = 0; i < m; i++) {
    for (let j = 0; j < n; j++) {
      const row = dp[i + 1];
      const previousRow = dp[i];
      if (!row || !previousRow) {
        continue;
      }
      row[j + 1] =
        beforeKeys[i] === afterKeys[j]
          ? (previousRow[j] ?? 0) + 1
          : Math.max(row[j] ?? 0, previousRow[j + 1] ?? 0);
    }
  }

  const reversed: DiffRun[] = [];
  let i = m;
  let j = n;
  while (i > 0 && j > 0) {
    if (beforeKeys[i - 1] === afterKeys[j - 1]) {
      reversed.push({
        type: "equal",
        before: before[i - 1] ?? "",
        after: after[j - 1] ?? "",
        units: 1,
      });
      i--;
      j--;
      continue;
    }
    // Backtracking emits in reverse, so pushing `ins` first here makes `del`
    // come BEFORE `ins` in the final left-to-right output. That ordering
    // matters at apply time: the inserted text lands after the
    // deletion-marked span, matching reader convention (strike-through → new
    // text) and the engine's existing expectation ("shallmust", not
    // "mustshall").
    if ((dp[i - 1]?.[j] ?? 0) > (dp[i]?.[j - 1] ?? 0)) {
      reversed.push({ type: "del", text: before[i - 1] ?? "" });
      i--;
    } else {
      reversed.push({ type: "ins", text: after[j - 1] ?? "" });
      j--;
    }
  }
  while (i > 0) {
    reversed.push({ type: "del", text: before[i - 1] ?? "" });
    i--;
  }
  while (j > 0) {
    reversed.push({ type: "ins", text: after[j - 1] ?? "" });
    j--;
  }

  const runs: DiffRun[] = [];
  for (const run of reversed.toReversed()) {
    const last = runs.at(-1);
    if (last?.type === "equal" && run.type === "equal") {
      last.before += run.before;
      last.after += run.after;
      last.units += run.units;
      continue;
    }
    if (
      (last?.type === "del" && run.type === "del") ||
      (last?.type === "ins" && run.type === "ins")
    ) {
      last.text += run.text;
      continue;
    }
    runs.push({ ...run });
  }
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
  const granularity = options.granularity ?? "word";
  const normalization = options.normalization ?? {};
  const beforeTokens = tokenize(before, granularity);
  const afterTokens = tokenize(after, granularity);
  if (beforeTokens.length === 0 && afterTokens.length === 0) {
    return [];
  }
  if (beforeTokens.length * afterTokens.length > MAX_WORD_DIFF_CELLS) {
    return wholeStringReplacement(before, after);
  }

  const aligned = alignTokens({ before: beforeTokens, after: afterTokens, normalization });
  const surviving = demoteRejectedMatches(aligned);
  if (isTooFragmented(surviving, (before.length + after.length) / 2)) {
    return wholeStringReplacement(before, after);
  }
  return toSegments(orderDeletionsFirst(surviving));
};
