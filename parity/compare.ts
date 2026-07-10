/**
 * Word rendering parity engine: pure comparison of two `DocGeom` values into
 * a `ParityResult`. No I/O, no browser: this module only manipulates the
 * already-extracted geometry, which keeps it fully unit-testable.
 *
 * The core problem is a sequence-alignment problem: match lines by text
 * across two documents (word = ground truth, folio = candidate), then diff
 * the matched pairs geometrically. See the module-level comments below for
 * the alignment algorithm and the gap-reconciliation (line-break) heuristic.
 */

import { DEFAULT_TOLERANCES } from "./config";
import { normalizeLineText, textSimilarity } from "./textNorm";
import type { ComparisonTolerances, Divergence, DocGeom, LineBox, ParityResult } from "./types";

/** Minimum text similarity for two lines to be considered the same line
 * (rather than a gap on one or both sides). */
const MATCH_THRESHOLD = 0.8;
/** Similarity threshold for reconciling a gap region as a line-break: the
 * concatenation of the unmatched word texts must read the same as the
 * concatenation of the unmatched folio texts. */
const CONCAT_THRESHOLD = 0.95;
/** Cost of leaving a line unmatched (gap) in the alignment DP. Small and
 * negative so two low-but-real similarities beat leaving both lines as gaps,
 * while still being cheap enough that a bad pairing (below MATCH_THRESHOLD,
 * scored -Infinity) is never preferred over a gap. */
const GAP_PENALTY = -0.2;
/** Full O(n*m) Needleman-Wunsch DP is used up to this many cells. Past it we
 * fall back to a cheaper per-page-window alignment (see `alignWindowed`). */
const MAX_DP_CELLS = 4_000_000;

/** A line flattened out of its page, keeping the page number it lives on. */
type FlatLine = { line: LineBox; page: number };

/** One step of a sequence alignment between a word-side and a folio-side
 * flattened line sequence. Indices are into the respective `FlatLine[]`. */
type AlignOp =
  | { kind: "match"; wordIdx: number; folioIdx: number }
  | { kind: "word-gap"; wordIdx: number }
  | { kind: "folio-gap"; folioIdx: number };

/** An alignment op stream reduced to its semantic outcome: matched pairs,
 * reconciled line-breaks, and genuinely missing/extra lines. */
type ResolvedItem =
  | { kind: "match"; wordIdx: number; folioIdx: number }
  | { kind: "line-break"; wordIdxs: number[]; folioIdxs: number[] }
  | { kind: "missing"; wordIdx: number }
  | { kind: "extra"; folioIdx: number };

export const compareGeoms = (
  word: DocGeom,
  folio: DocGeom,
  tolerances: ComparisonTolerances = DEFAULT_TOLERANCES,
): ParityResult => {
  const wordFlat = flatten(word);
  const folioFlat = flatten(folio);

  const divergences: Divergence[] = [];
  if (word.pages.length !== folio.pages.length) {
    divergences.push({ kind: "page-count", word: word.pages.length, folio: folio.pages.length });
  }

  const cellCount = wordFlat.length * folioFlat.length;
  const resolved =
    cellCount > MAX_DP_CELLS
      ? alignWindowed({ word, folio, wordFlat, folioFlat })
      : resolveOps(alignFull(wordFlat, folioFlat), wordFlat, folioFlat);

  const matches = resolved.filter(
    (item): item is Extract<ResolvedItem, { kind: "match" }> => item.kind === "match",
  );
  const medianYOffsetsByPageRegion = pageRegionMedianYOffsets(matches, wordFlat, folioFlat);
  const medianYOffsetPt = median([...medianYOffsetsByPageRegion.values()]);

  const { orderedDivergences, matchedGeomPass } = diffMatches({
    resolved,
    wordFlat,
    folioFlat,
    tolerances,
    medianYOffsetsByPageRegion,
  });
  divergences.push(...orderedDivergences);

  const totalWordLines = wordFlat.length;
  const score =
    totalWordLines === 0
      ? scoreForEmptyWordDoc(folioFlat.length)
      : matchedGeomPass / totalWordLines;

  return {
    file: word.file,
    score,
    wordPages: word.pages.length,
    folioPages: folio.pages.length,
    totalWordLines,
    matchedLines: matches.length,
    medianYOffsetPt,
    divergences,
  };
};

/** Score when the Word side has no lines at all: 1 if folio is also empty
 * (nothing to diverge on), 0 otherwise (every folio line is unaccounted for). */
const scoreForEmptyWordDoc = (folioLineCount: number): number => (folioLineCount === 0 ? 1 : 0);

const stripSpaces = (text: string): string => text.replaceAll(" ", "");

/** Minimum vertical-overlap ratio (of the shorter box) for two boxes to count
 * as the same visual row. */
const ROW_OVERLAP_RATIO = 0.5;
/** Maximum horizontal gap (pt) between same-row boxes that still merges them.
 * Word emits list markers as separate ink boxes ~10pt left of the item text,
 * and tabbed legal clauses can leave ~23pt between the marker and text; table
 * cells sit farther apart (>25pt) and must stay separate. */
const ROW_MERGE_GAP_PT = 24;
const MARKER_ROW_MERGE_GAP_PT = 36;

/** Clusters boxes into visual rows (>= ROW_OVERLAP_RATIO vertical overlap),
 * then merges row neighbours within ROW_MERGE_GAP_PT of each other into a
 * single LineBox — bullet/number markers join their item text the way folio
 * paints them, while table cells (large gaps) stay separate boxes. The result
 * is ordered row-by-row, left-to-right within a row: Word's ink boxes on one
 * visual row can differ by fractions of a pt vertically (e.g. table cells), so
 * a raw (y, x) sort would order the same row differently on the two sides and
 * derail the sequence alignment. Applied to BOTH sides so the pass itself
 * can't create asymmetry. */
export const mergeVisualRows = (lines: LineBox[]): LineBox[] => {
  const rows: LineBox[][] = [];
  for (const line of [...lines].sort((a, b) => a.yPt - b.yPt || a.xPt - b.xPt)) {
    const row = rows.at(-1);
    if (row && row.some((member) => isSameVisualRow(member, line))) {
      row.push(line);
      continue;
    }
    rows.push([line]);
  }

  const merged: LineBox[] = [];
  for (const row of rows) {
    row.sort((a, b) => a.xPt - b.xPt);
    let current = row[0];
    if (!current) continue;
    for (const next of row.slice(1)) {
      if (shouldMergeRowBoxes(current, next)) {
        current = mergeBoxes(current, next);
        continue;
      }
      merged.push(current);
      current = next;
    }
    merged.push(current);
  }
  return merged;
};

const shouldMergeRowBoxes = (current: LineBox, next: LineBox): boolean => {
  if (current.visualGroup !== next.visualGroup) {
    return false;
  }
  const gap = horizontalGap(current, next);
  if (gap <= ROW_MERGE_GAP_PT) {
    return true;
  }
  return isStandaloneListMarker(current.text) && gap <= MARKER_ROW_MERGE_GAP_PT;
};

const isStandaloneListMarker = (text: string): boolean => {
  const normalized = normalizeLineText(text);
  return /^((\([A-Za-z0-9ivxlcdm]+\))|([A-Za-z0-9ivxlcdm]+[.)])|[•·])$/iu.test(normalized);
};

const isSameVisualRow = (a: LineBox, b: LineBox): boolean => {
  const overlap = Math.min(a.yPt + a.heightPt, b.yPt + b.heightPt) - Math.max(a.yPt, b.yPt);
  const minHeight = Math.min(a.heightPt, b.heightPt);
  if (minHeight <= 0) return false;
  return overlap / minHeight >= ROW_OVERLAP_RATIO;
};

const horizontalGap = (a: LineBox, b: LineBox): number => {
  // Sorted by (y, x) the gap is usually b.left - a.right, but keep it
  // symmetric for overlapping boxes (gap <= 0 merges).
  return Math.max(a.xPt, b.xPt) - Math.min(a.xPt + a.widthPt, b.xPt + b.widthPt);
};

const mergeBoxes = (a: LineBox, b: LineBox): LineBox => {
  const [left, right] = a.xPt <= b.xPt ? [a, b] : [b, a];
  const xPt = Math.min(a.xPt, b.xPt);
  const yPt = Math.min(a.yPt, b.yPt);
  const text = `${left.text} ${right.text}`;
  // Prefer the left box's font, but fall back to the right box's so font
  // metadata is not discarded when only the right side carries it.
  const fontName = left.fontName ?? right.fontName;
  const fontSizePt = left.fontSizePt ?? right.fontSizePt;
  return {
    text,
    normText: normalizeLineText(text),
    xPt,
    yPt,
    widthPt: Math.max(a.xPt + a.widthPt, b.xPt + b.widthPt) - xPt,
    heightPt: Math.max(a.yPt + a.heightPt, b.yPt + b.heightPt) - yPt,
    region: left.region,
    ...(left.visualGroup !== undefined ? { visualGroup: left.visualGroup } : {}),
    ...(fontName !== undefined ? { fontName } : {}),
    ...(fontSizePt !== undefined ? { fontSizePt } : {}),
  };
};

/** Flattens a doc's pages into an ordered line sequence (merging visual rows
 * first), dropping lines with empty normalized text defensively (extractors
 * already drop these). */
const flatten = (doc: DocGeom): FlatLine[] => {
  const flat: FlatLine[] = [];
  for (const page of doc.pages) {
    for (const line of mergeVisualRows(page.lines)) {
      if (line.normText.length === 0) continue;
      flat.push({ line, page: page.number });
    }
  }
  return flat;
};

/**
 * Needleman-Wunsch global alignment of two flattened line sequences on
 * `normText` similarity. Diagonal (match) moves are only available when
 * `textSimilarity >= MATCH_THRESHOLD`; otherwise the cell falls back to a
 * gap (word-only or folio-only), so low-similarity pairs are never forced
 * into a "match" that would then need a text-mismatch divergence.
 *
 * O(n*m) time and space; callers must guard the size (see MAX_DP_CELLS).
 */
const alignFull = (wordFlat: FlatLine[], folioFlat: FlatLine[]): AlignOp[] => {
  const n = wordFlat.length;
  const m = folioFlat.length;
  const cols = m + 1;
  const score = new Float64Array((n + 1) * cols);
  // 0 = diagonal (match), 1 = up (word-gap), 2 = left (folio-gap).
  const trace = new Uint8Array((n + 1) * cols);

  for (let i = 1; i <= n; i++) {
    score[i * cols] = (score[(i - 1) * cols] ?? 0) + GAP_PENALTY;
    trace[i * cols] = 1;
  }
  for (let j = 1; j <= m; j++) {
    score[j] = (score[j - 1] ?? 0) + GAP_PENALTY;
    trace[j] = 2;
  }

  for (let i = 1; i <= n; i++) {
    const wordText = wordFlat[i - 1]?.line.normText ?? "";
    for (let j = 1; j <= m; j++) {
      const folioText = folioFlat[j - 1]?.line.normText ?? "";
      const sim = textSimilarity(wordText, folioText);
      const diagScore =
        sim >= MATCH_THRESHOLD ? (score[(i - 1) * cols + (j - 1)] ?? 0) + sim : -Infinity;
      const upScore = (score[(i - 1) * cols + j] ?? 0) + GAP_PENALTY;
      const leftScore = (score[i * cols + (j - 1)] ?? 0) + GAP_PENALTY;

      let best = diagScore;
      let move = 0;
      if (upScore > best) {
        best = upScore;
        move = 1;
      }
      if (leftScore > best) {
        best = leftScore;
        move = 2;
      }

      score[i * cols + j] = best;
      trace[i * cols + j] = move;
    }
  }

  const ops: AlignOp[] = [];
  let i = n;
  let j = m;
  while (i > 0 || j > 0) {
    const move = trace[i * cols + j];
    if (move === 0) {
      ops.push({ kind: "match", wordIdx: i - 1, folioIdx: j - 1 });
      i--;
      j--;
    } else if (move === 1) {
      ops.push({ kind: "word-gap", wordIdx: i - 1 });
      i--;
    } else {
      ops.push({ kind: "folio-gap", folioIdx: j - 1 });
      j--;
    }
  }
  ops.reverse();
  return ops;
};

/**
 * Fallback alignment for pathologically large documents where the full
 * O(n*m) DP would be too large (see MAX_DP_CELLS). Aligns each word page's
 * lines against the folio lines of the surrounding window (that word page's
 * position, plus one neighbour on each side, clamped to the folio page
 * range), consuming folio lines at most once so later windows cannot rematch
 * an already-matched line.
 *
 * Line-break reconciliation happens per-window (a gap region spanning a
 * word-gap and a still-available folio candidate resolves the same way it
 * would in `alignFull`). A folio line that stays unconsumed through every
 * window it was eligible for is only finalized as an "extra" line once all
 * word pages have been processed, so leftovers from this path are appended
 * once at the end rather than interleaved at their true document position.
 * This is a deliberate approximation for an escape hatch that only exists to
 * bound pathological input sizes; it is not the common path.
 */
type AlignWindowedOptions = {
  word: DocGeom;
  folio: DocGeom;
  wordFlat: FlatLine[];
  folioFlat: FlatLine[];
};

const alignWindowed = ({
  word,
  folio,
  wordFlat,
  folioFlat,
}: AlignWindowedOptions): ResolvedItem[] => {
  const wordByPagePos = groupByPagePos(wordFlat, word);
  const folioByPagePos = groupByPagePos(folioFlat, folio);
  const consumed = Array.from({ length: folioFlat.length }, () => false);
  const resolved: ResolvedItem[] = [];

  for (let pk = 0; pk < word.pages.length; pk++) {
    const wordIdxs = wordByPagePos[pk] ?? [];
    if (wordIdxs.length === 0) continue;

    const lo = Math.max(0, pk - 1);
    const hi = Math.min(folio.pages.length - 1, pk + 1);
    const candidateIdxs: number[] = [];
    for (let p = lo; p <= hi; p++) {
      for (const idx of folioByPagePos[p] ?? []) {
        if (!consumed[idx]) candidateIdxs.push(idx);
      }
    }

    const localWord = wordIdxs
      .map((idx) => wordFlat[idx])
      .filter((fl): fl is FlatLine => fl !== undefined);
    const localFolio = candidateIdxs
      .map((idx) => folioFlat[idx])
      .filter((fl): fl is FlatLine => fl !== undefined);
    const localResolved = resolveOps(alignFull(localWord, localFolio), localWord, localFolio);

    for (const item of localResolved) {
      if (item.kind === "match") {
        const folioIdx = candidateIdxs[item.folioIdx];
        const wordIdx = wordIdxs[item.wordIdx];
        if (folioIdx === undefined || wordIdx === undefined) continue;
        consumed[folioIdx] = true;
        resolved.push({ kind: "match", wordIdx, folioIdx });
      } else if (item.kind === "line-break") {
        const globalFolioIdxs = item.folioIdxs
          .map((idx) => candidateIdxs[idx])
          .filter((idx): idx is number => idx !== undefined);
        const globalWordIdxs = item.wordIdxs
          .map((idx) => wordIdxs[idx])
          .filter((idx): idx is number => idx !== undefined);
        for (const idx of globalFolioIdxs) consumed[idx] = true;
        resolved.push({ kind: "line-break", wordIdxs: globalWordIdxs, folioIdxs: globalFolioIdxs });
      } else if (item.kind === "missing") {
        const wordIdx = wordIdxs[item.wordIdx];
        if (wordIdx !== undefined) resolved.push({ kind: "missing", wordIdx });
      }
      // "extra" items are left pending: the candidate folio line was not
      // consumed above, so it remains available for a neighbouring
      // window. Unconsumed leftovers are finalized below.
    }
  }

  for (let idx = 0; idx < folioFlat.length; idx++) {
    if (!consumed[idx]) resolved.push({ kind: "extra", folioIdx: idx });
  }

  return resolved;
};

/** Groups flat-line indices by the position (0-based) of their page within
 * `doc.pages`, for the windowed alignment fallback. */
const groupByPagePos = (flat: FlatLine[], doc: DocGeom): number[][] => {
  const pageNumberToPos = new Map<number, number>();
  doc.pages.forEach((page, pos) => pageNumberToPos.set(page.number, pos));
  const groups: number[][] = doc.pages.map(() => []);
  flat.forEach((flatLine, idx) => {
    const pos = pageNumberToPos.get(flatLine.page);
    if (pos !== undefined) groups[pos]?.push(idx);
  });
  return groups;
};

/**
 * Reduces an alignment op stream to matched pairs, reconciled line-breaks,
 * and genuinely missing/extra lines, preserving the op stream's order. A
 * "gap region" is a maximal run of consecutive word-gap/folio-gap ops
 * between two matches (or at either end of the sequence); see
 * `reconcileGap` for how a region is resolved.
 */
const resolveOps = (
  ops: AlignOp[],
  wordFlat: FlatLine[],
  folioFlat: FlatLine[],
): ResolvedItem[] => {
  const resolved: ResolvedItem[] = [];
  let gapWordIdxs: number[] = [];
  let gapFolioIdxs: number[] = [];

  const flushGap = () => {
    if (gapWordIdxs.length === 0 && gapFolioIdxs.length === 0) return;
    resolved.push(
      ...reconcileGap({ wordIdxs: gapWordIdxs, folioIdxs: gapFolioIdxs, wordFlat, folioFlat }),
    );
    gapWordIdxs = [];
    gapFolioIdxs = [];
  };

  for (const op of ops) {
    if (op.kind === "match") {
      flushGap();
      resolved.push({ kind: "match", wordIdx: op.wordIdx, folioIdx: op.folioIdx });
    } else if (op.kind === "word-gap") {
      gapWordIdxs.push(op.wordIdx);
    } else {
      gapFolioIdxs.push(op.folioIdx);
    }
  }
  flushGap();
  return resolved;
};

/**
 * Resolves one gap region. If both sides are non-empty and the
 * space-joined, re-normalized concatenation of the word texts reads the same
 * as the folio texts (similarity >= CONCAT_THRESHOLD), the whole region is a
 * single line-break (line-break-position difference), covering the common
 * 1-to-N and N-to-1 cases. Otherwise every word line is missing and every
 * folio line is extra.
 *
 * Greedy sub-splitting of large, only-partially-reconcilable regions (e.g.
 * matching prefixes progressively) is spec-optional and not implemented
 * here: a region that doesn't reconcile as a whole falls through entirely to
 * missing/extra rather than being partially salvaged.
 */
type ReconcileGapOptions = {
  wordIdxs: number[];
  folioIdxs: number[];
  wordFlat: FlatLine[];
  folioFlat: FlatLine[];
};

const reconcileGap = ({
  wordIdxs,
  folioIdxs,
  wordFlat,
  folioFlat,
}: ReconcileGapOptions): ResolvedItem[] => {
  if (wordIdxs.length > 0 && folioIdxs.length > 0) {
    const wordConcat = normalizeLineText(
      wordIdxs.map((idx) => wordFlat[idx]?.line.normText ?? "").join(" "),
    );
    const folioConcat = normalizeLineText(
      folioIdxs.map((idx) => folioFlat[idx]?.line.normText ?? "").join(" "),
    );
    if (
      wordIdxs.length === 1 &&
      folioIdxs.length === 1 &&
      stripSpaces(wordConcat) === stripSpaces(folioConcat)
    ) {
      const wordIdx = wordIdxs[0];
      const folioIdx = folioIdxs[0];
      if (wordIdx !== undefined && folioIdx !== undefined) {
        return [{ kind: "match", wordIdx, folioIdx }];
      }
    }
    if (textSimilarity(wordConcat, folioConcat) >= CONCAT_THRESHOLD) {
      return [{ kind: "line-break", wordIdxs, folioIdxs }];
    }
  }
  return [
    ...wordIdxs.map((wordIdx): ResolvedItem => ({ kind: "missing", wordIdx })),
    ...folioIdxs.map((folioIdx): ResolvedItem => ({ kind: "extra", folioIdx })),
  ];
};

const pageRegionKey = (page: number, region: LineBox["region"]): string =>
  `${page}:${region ?? "unknown"}`;

const matchedPageRegionKey = (word: FlatLine, folio: FlatLine): string =>
  pageRegionKey(word.page, folio.line.region);

const pageRegionMedianYOffsets = (
  matches: Extract<ResolvedItem, { kind: "match" }>[],
  wordFlat: FlatLine[],
  folioFlat: FlatLine[],
): Map<string, number> => {
  const deltasByPageRegion = new Map<string, number[]>();
  for (const m of matches) {
    const w = wordFlat[m.wordIdx];
    const f = folioFlat[m.folioIdx];
    if (!w || !f || w.page !== f.page) continue;
    const key = matchedPageRegionKey(w, f);
    const deltas = deltasByPageRegion.get(key) ?? [];
    deltas.push(f.line.yPt - w.line.yPt);
    deltasByPageRegion.set(key, deltas);
  }
  return new Map([...deltasByPageRegion].map(([key, deltas]) => [key, median(deltas)]));
};

const median = (values: number[]): number => {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 0) return ((sorted[mid - 1] ?? 0) + (sorted[mid] ?? 0)) / 2;
  return sorted[mid] ?? 0;
};

/** Result of walking the resolved items into divergences: the divergences
 * in document order, plus the count of matched pairs that pass every
 * geometric check (used for the parity score numerator). */
type DiffMatchesResult = { orderedDivergences: Divergence[]; matchedGeomPass: number };
type DiffMatchesOptions = {
  resolved: ResolvedItem[];
  wordFlat: FlatLine[];
  folioFlat: FlatLine[];
  tolerances: ComparisonTolerances;
  medianYOffsetsByPageRegion: ReadonlyMap<string, number>;
};

const diffMatches = ({
  resolved,
  wordFlat,
  folioFlat,
  tolerances,
  medianYOffsetsByPageRegion,
}: DiffMatchesOptions): DiffMatchesResult => {
  const orderedDivergences: Divergence[] = [];
  let matchedGeomPass = 0;

  for (const item of resolved) {
    if (item.kind === "match") {
      const w = wordFlat[item.wordIdx];
      const f = folioFlat[item.folioIdx];
      if (!w || !f) continue;
      const samePage = w.page === f.page;

      // Space-insensitive: visual-row merging joins Word's separate marker
      // boxes with a space while folio's painter inlines markers without one,
      // and that spacing difference is not a text fidelity issue.
      if (stripSpaces(w.line.normText) !== stripSpaces(f.line.normText)) {
        orderedDivergences.push({
          kind: "text-mismatch",
          page: w.page,
          wordText: w.line.normText,
          folioText: f.line.normText,
        });
      }
      if (!samePage) {
        orderedDivergences.push({
          kind: "pagination",
          text: w.line.normText,
          wordPage: w.page,
          folioPage: f.page,
        });
      }

      const xDelta = checkXDrift(w.line, f.line, tolerances);
      if (xDelta !== null) {
        orderedDivergences.push({
          kind: "x-drift",
          page: w.page,
          text: w.line.normText,
          deltaPt: xDelta,
        });
      }
      const widthDelta = checkWidthDrift(w.line, f.line, tolerances);
      if (widthDelta !== null) {
        orderedDivergences.push({
          kind: "width-drift",
          page: w.page,
          text: w.line.normText,
          deltaPt: widthDelta,
        });
      }

      let yDelta: number | null = null;
      if (samePage) {
        yDelta = checkYDrift(
          w.line,
          f.line,
          medianYOffsetsByPageRegion.get(matchedPageRegionKey(w, f)) ?? 0,
          tolerances,
        );
        if (yDelta !== null) {
          orderedDivergences.push({
            kind: "y-drift",
            page: w.page,
            text: w.line.normText,
            residualPt: yDelta,
          });
        }
      }

      if (samePage && xDelta === null && widthDelta === null && yDelta === null) matchedGeomPass++;
    } else if (item.kind === "line-break") {
      const firstWordIdx = item.wordIdxs[0];
      const page = firstWordIdx === undefined ? undefined : wordFlat[firstWordIdx]?.page;
      if (page !== undefined) {
        orderedDivergences.push({
          kind: "line-break",
          page,
          wordTexts: item.wordIdxs.map((idx) => wordFlat[idx]?.line.normText ?? ""),
          folioTexts: item.folioIdxs.map((idx) => folioFlat[idx]?.line.normText ?? ""),
        });
      }
    } else if (item.kind === "missing") {
      const w = wordFlat[item.wordIdx];
      if (w) orderedDivergences.push({ kind: "missing-line", page: w.page, text: w.line.normText });
    } else {
      const f = folioFlat[item.folioIdx];
      if (f) orderedDivergences.push({ kind: "extra-line", page: f.page, text: f.line.normText });
    }
  }

  return { orderedDivergences, matchedGeomPass };
};

/** x-drift: folio's left edge vs word's, page-relative pt. */
const checkXDrift = (
  wordLine: LineBox,
  folioLine: LineBox,
  tolerances: ComparisonTolerances,
): number | null => {
  const delta = folioLine.xPt - wordLine.xPt;
  return Math.abs(delta) > tolerances.xPt ? delta : null;
};

/** width-drift: passes if within the absolute tolerance OR the relative
 * tolerance (a fraction of the word line's own width). */
const checkWidthDrift = (
  wordLine: LineBox,
  folioLine: LineBox,
  tolerances: ComparisonTolerances,
): number | null => {
  const delta = folioLine.widthPt - wordLine.widthPt;
  const withinAbsolute = Math.abs(delta) <= tolerances.widthPt;
  const withinRelative = Math.abs(delta) <= tolerances.widthRelative * wordLine.widthPt;
  return withinAbsolute || withinRelative ? null : delta;
};

/**
 * y-drift residual: word-side boxes are glyph-ink bounds and folio-side
 * boxes are line-height boxes, so a constant vertical offset between the two
 * extractors is expected. `medianOffsetPt` (the median of folio-minus-word y
 * across every same-page matched pair) absorbs that constant; only the
 * residual after subtracting it is a real divergence.
 */
const checkYDrift = (
  wordLine: LineBox,
  folioLine: LineBox,
  medianOffsetPt: number,
  tolerances: ComparisonTolerances,
): number | null => {
  const residual = folioLine.yPt - wordLine.yPt - medianOffsetPt;
  return Math.abs(residual) > tolerances.yResidualPt ? residual : null;
};
