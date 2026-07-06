/**
 * Document version-diff engine: compare two `.docx` buffers block by block
 * and produce a structured, LLM-summarizable diff.
 *
 * Both buffers are parsed through {@link FolioDocxReviewer} — the same
 * headless parsing + clean-text path `read_document` / `read_changes` use —
 * so the comparison runs over each document's AS-ACCEPTED view: any pending
 * tracked changes already present in EITHER buffer count as applied before
 * the two are compared. Two documents that agree once their own pending
 * redlines are accepted report as unchanged, even if the underlying
 * tracked-change history differs.
 *
 * ## Alignment
 *
 * Blocks are paired across the two snapshots in three passes, each only
 * considering blocks the previous pass left unpaired:
 *
 * 1. **Stable-id pairing.** Blocks whose ids are equal AND not a `seq-NNNN`
 *    positional fallback ({@link getFolioParaIdFromBlockId} returns non-null)
 *    are paired directly. This covers both id shapes a snapshot can carry:
 *    - A real Word `w14:paraId`: stable identity, independent of text — an
 *      equal-id pair with different text is a genuine edit (`modified`).
 *    - `FolioDocxReviewer`'s deterministic fallback id (assigned when the
 *      source paragraph has no `w14:paraId`), which hashes the paragraph's
 *      TEXT plus its document ordinal. It is structurally indistinguishable
 *      from a real paraId ({@link getFolioParaIdFromBlockId} can't tell them
 *      apart), but pairing on equality is still safe: two equal deterministic
 *      ids necessarily came from identical text at an identical ordinal, so
 *      the pair is always text-equal (`unchanged`) — never a false
 *      `modified`. What it can't do is FIND a paragraph whose ordinal shifted
 *      (an insertion/deletion earlier in the document) even though its text
 *      is unchanged: that pair has two different fallback ids and falls
 *      through to pass 2.
 * 2. **Exact-text pairing.** An order-preserving LCS over remaining blocks,
 *    matched by exact text equality. This is what recovers same-text blocks
 *    that pass 1 missed because a fallback id shifted with the ordinal. Its
 *    O(m·n) table is skipped ({@link exceedsLcsBudget}) once the unpaired
 *    counts on both sides would exceed a fixed cell budget, so a document
 *    with few/no stable ids can't force a quadratic-sized allocation; those
 *    blocks fall through to pass 3 instead.
 * 3. **Positional fallback.** Whatever a monotonicity filter leaves
 *    unpaired is split into the gaps between anchored pairs (pass 1 + 2,
 *    time-ordered); within each gap the shorter side is zipped positionally
 *    against the longer one (`modified`), and any excess on either side is
 *    reported as `added` / `deleted`.
 *
 * The combined anchor set from passes 1 and 2 is re-filtered to the longest
 * increasing subsequence by revised-side index before pass 3 runs, so a
 * pathological crossing match (content reordered across versions) can't
 * produce an out-of-order gap — the alignment always walks both documents
 * forward.
 */

import { diffWordSegments } from "@stll/folio-core/ai-edits";
import type { WordDiffSegment } from "@stll/folio-core/ai-edits";
import { FolioDocxReviewer, getFolioParaIdFromBlockId } from "@stll/folio-core/server";
import type { FolioAIBlock } from "@stll/folio-core/server";

/** One word-level diff segment within a `modified` block. Mirrors {@link WordDiffSegment}. */
export type FolioAgentVersionDiffSegment = WordDiffSegment;

/** One block-level change between two document versions, in revised-side document order. */
export type FolioAgentBlockDiff =
  | { type: "added"; blockId: string; kind: string; text: string }
  | { type: "deleted"; blockId: string; kind: string; text: string }
  | { type: "modified"; blockId: string; kind: string; segments: FolioAgentVersionDiffSegment[] };

/** Result of {@link compareDocxVersions}. */
export type FolioAgentVersionDiff = {
  /** Every added, deleted, or modified block, in revised-side document order (deletions slotted where they sat). */
  changes: FolioAgentBlockDiff[];
  /** Counts across every paired/unpaired block, including the unchanged blocks `changes` omits. */
  summaryCounts: { added: number; deleted: number; modified: number; unchanged: number };
};

type BlockPair = { baseIndex: number; revisedIndex: number };
type IndexedBlock = { block: FolioAIBlock; index: number };

const isStableBlockId = (id: string): boolean => getFolioParaIdFromBlockId(id) !== null;

const index = (blocks: readonly FolioAIBlock[]): IndexedBlock[] =>
  blocks.map((block, blockIndex) => ({ block, index: blockIndex }));

/**
 * Longest increasing subsequence by `revisedIndex`, assuming `pairs` is
 * already sorted by `baseIndex` ascending. Drops any pair that would make
 * the alignment walk backward in the revised document — the guard against
 * both id collisions (pass 1) and any crossing match (pass 1 + 2 combined).
 */
const longestIncreasingByRevisedIndex = (pairs: readonly BlockPair[]): BlockPair[] => {
  if (pairs.length === 0) {
    return [];
  }
  const lengths = Array.from<number>({ length: pairs.length }).fill(1);
  const predecessors = Array.from<number>({ length: pairs.length }).fill(-1);
  let bestEnd = 0;
  for (let i = 0; i < pairs.length; i++) {
    for (let j = 0; j < i; j++) {
      const current = pairs[j];
      const candidate = pairs[i];
      if (!current || !candidate) {
        continue;
      }
      if (current.revisedIndex < candidate.revisedIndex && (lengths[j] ?? 0) + 1 > (lengths[i] ?? 0)) {
        lengths[i] = (lengths[j] ?? 0) + 1;
        predecessors[i] = j;
      }
    }
    if ((lengths[i] ?? 0) > (lengths[bestEnd] ?? 0)) {
      bestEnd = i;
    }
  }
  const ordered: BlockPair[] = [];
  for (let cursor = bestEnd; cursor !== -1; cursor = predecessors[cursor] ?? -1) {
    const pair = pairs[cursor];
    if (pair) {
      ordered.push(pair);
    }
  }
  return ordered.toReversed();
};

/** Pass 1: pair blocks with equal, non-`seq-NNNN` ids. See the module doc comment. */
const pairByStableId = (base: readonly FolioAIBlock[], revised: readonly FolioAIBlock[]): BlockPair[] => {
  const revisedIndexById = new Map<string, number>();
  revised.forEach((block, revisedIndex) => {
    if (isStableBlockId(block.id)) {
      revisedIndexById.set(block.id, revisedIndex);
    }
  });

  const candidates: BlockPair[] = [];
  base.forEach((block, baseIndex) => {
    if (!isStableBlockId(block.id)) {
      return;
    }
    const revisedIndex = revisedIndexById.get(block.id);
    if (revisedIndex !== undefined) {
      candidates.push({ baseIndex, revisedIndex });
    }
  });
  return longestIncreasingByRevisedIndex(candidates);
};

/**
 * Cell budget for pass 2's O(m·n) exact-text LCS table (`dp` below allocates
 * `(m + 1) * (n + 1)` numbers). A document with no `w14:paraId`s — or an
 * adversarial one crafted to defeat pass 1 — can leave thousands of blocks
 * unpaired on both sides; without a cap, `pairByExactText` would allocate a
 * quadratic-sized table for it. Past this budget, {@link exceedsLcsBudget}
 * makes pass 2 back off entirely so alignment falls through to pass 3's
 * linear positional zip instead — pairing is less precise for these
 * degenerate inputs, but memory use stays bounded.
 */
const MAX_LCS_CELLS = 4_000_000;

/** True when an `unpairedBaseCount * unpairedRevisedCount` LCS table would exceed {@link MAX_LCS_CELLS}. */
export const exceedsLcsBudget = (unpairedBaseCount: number, unpairedRevisedCount: number): boolean =>
  unpairedBaseCount * unpairedRevisedCount > MAX_LCS_CELLS;

/** Pass 2: order-preserving LCS by exact text equality over the blocks pass 1 left unpaired. */
const pairByExactText = (base: readonly IndexedBlock[], revised: readonly IndexedBlock[]): BlockPair[] => {
  const m = base.length;
  const n = revised.length;
  if (m === 0 || n === 0) {
    return [];
  }
  if (exceedsLcsBudget(m, n)) {
    return [];
  }
  const dp: number[][] = Array.from({ length: m + 1 }, () =>
    Array.from<number>({ length: n + 1 }).fill(0),
  );
  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      const row = dp[i];
      const nextRow = dp[i + 1];
      if (!row || !nextRow) {
        continue;
      }
      const baseText = base[i]?.block.text;
      const revisedText = revised[j]?.block.text;
      row[j] =
        baseText === revisedText
          ? (nextRow[j + 1] ?? 0) + 1
          : Math.max(nextRow[j] ?? 0, row[j + 1] ?? 0);
    }
  }

  const pairs: BlockPair[] = [];
  let i = 0;
  let j = 0;
  while (i < m && j < n) {
    const baseEntry = base[i];
    const revisedEntry = revised[j];
    if (!baseEntry || !revisedEntry) {
      break;
    }
    if (baseEntry.block.text === revisedEntry.block.text) {
      pairs.push({ baseIndex: baseEntry.index, revisedIndex: revisedEntry.index });
      i++;
      j++;
      continue;
    }
    const down = dp[i + 1]?.[j] ?? 0;
    const right = dp[i]?.[j + 1] ?? 0;
    if (down >= right) {
      i++;
    } else {
      j++;
    }
  }
  return pairs;
};

const ELLIPSIS = "…";
const UNCHANGED_EDGE_CHARS = 30;

/** Collapse a long `equal`-type run to ~30 chars on each side with an ellipsis in between. */
const truncateUnchangedRun = (text: string): string => {
  if (text.length <= UNCHANGED_EDGE_CHARS * 2 + ELLIPSIS.length) {
    return text;
  }
  return `${text.slice(0, UNCHANGED_EDGE_CHARS)}${ELLIPSIS}${text.slice(-UNCHANGED_EDGE_CHARS)}`;
};

/** Render one word-diff segment: plain for `equal` (truncated if long), bracketed for `del` / `ins`. */
const renderSegment = (segment: FolioAgentVersionDiffSegment): string => {
  switch (segment.type) {
    case "equal":
      return truncateUnchangedRun(segment.text);
    case "del":
      return `[-${segment.text}-]`;
    case "ins":
      return `{+${segment.text}+}`;
    default:
      return "";
  }
};

const formatChangeLine = (change: FolioAgentBlockDiff): string => {
  if (change.type === "added") {
    return `+ [${change.blockId}] added: ${change.text}`;
  }
  if (change.type === "deleted") {
    return `- [${change.blockId}] deleted: ${change.text}`;
  }
  return `~ [${change.blockId}] modified: ${change.segments.map(renderSegment).join("")}`;
};

/**
 * Compare two `.docx` buffers and return a structured, block-level diff.
 * See the module doc comment for the as-accepted comparison semantics and
 * the three-pass alignment algorithm.
 */
export const compareDocxVersions = async (
  base: ArrayBuffer,
  revised: ArrayBuffer,
): Promise<FolioAgentVersionDiff> => {
  const [baseReviewer, revisedReviewer] = await Promise.all([
    FolioDocxReviewer.fromBuffer(base),
    FolioDocxReviewer.fromBuffer(revised),
  ]);
  const baseBlocks = baseReviewer.snapshot().blocks;
  const revisedBlocks = revisedReviewer.snapshot().blocks;

  const stableIdAnchors = pairByStableId(baseBlocks, revisedBlocks);
  const usedBaseIndexes = new Set(stableIdAnchors.map((anchor) => anchor.baseIndex));
  const usedRevisedIndexes = new Set(stableIdAnchors.map((anchor) => anchor.revisedIndex));

  const baseRemaining = index(baseBlocks).filter(({ index: i }) => !usedBaseIndexes.has(i));
  const revisedRemaining = index(revisedBlocks).filter(({ index: i }) => !usedRevisedIndexes.has(i));
  const exactTextAnchors = pairByExactText(baseRemaining, revisedRemaining);

  const anchors = longestIncreasingByRevisedIndex(
    [...stableIdAnchors, ...exactTextAnchors].toSorted((a, b) => a.baseIndex - b.baseIndex),
  );

  const changes: FolioAgentBlockDiff[] = [];
  const counts = { added: 0, deleted: 0, modified: 0, unchanged: 0 };

  /** Pass 3: positionally zip the leftover blocks in one gap between anchors. */
  const emitGap = (baseFrom: number, baseTo: number, revisedFrom: number, revisedTo: number): void => {
    const baseSlice = baseBlocks.slice(baseFrom, baseTo);
    const revisedSlice = revisedBlocks.slice(revisedFrom, revisedTo);
    const pairedCount = Math.min(baseSlice.length, revisedSlice.length);

    for (let k = 0; k < pairedCount; k++) {
      const baseBlock = baseSlice[k];
      const revisedBlock = revisedSlice[k];
      if (!baseBlock || !revisedBlock) {
        continue;
      }
      if (baseBlock.text === revisedBlock.text) {
        counts.unchanged++;
        continue;
      }
      counts.modified++;
      changes.push({
        type: "modified",
        blockId: revisedBlock.id,
        kind: revisedBlock.kind,
        segments: diffWordSegments(baseBlock.text, revisedBlock.text),
      });
    }
    for (let k = pairedCount; k < baseSlice.length; k++) {
      const baseBlock = baseSlice[k];
      if (!baseBlock) {
        continue;
      }
      counts.deleted++;
      changes.push({ type: "deleted", blockId: baseBlock.id, kind: baseBlock.kind, text: baseBlock.text });
    }
    for (let k = pairedCount; k < revisedSlice.length; k++) {
      const revisedBlock = revisedSlice[k];
      if (!revisedBlock) {
        continue;
      }
      counts.added++;
      changes.push({ type: "added", blockId: revisedBlock.id, kind: revisedBlock.kind, text: revisedBlock.text });
    }
  };

  let baseCursor = 0;
  let revisedCursor = 0;
  for (const anchor of anchors) {
    emitGap(baseCursor, anchor.baseIndex, revisedCursor, anchor.revisedIndex);

    const baseBlock = baseBlocks[anchor.baseIndex];
    const revisedBlock = revisedBlocks[anchor.revisedIndex];
    if (baseBlock && revisedBlock) {
      if (baseBlock.text === revisedBlock.text) {
        counts.unchanged++;
      } else {
        counts.modified++;
        changes.push({
          type: "modified",
          blockId: revisedBlock.id,
          kind: revisedBlock.kind,
          segments: diffWordSegments(baseBlock.text, revisedBlock.text),
        });
      }
    }

    baseCursor = anchor.baseIndex + 1;
    revisedCursor = anchor.revisedIndex + 1;
  }
  emitGap(baseCursor, baseBlocks.length, revisedCursor, revisedBlocks.length);

  return { changes, summaryCounts: counts };
};

/**
 * Render a {@link FolioAgentVersionDiff} as compact, deterministic text for a
 * model prompt: a header line with the summary counts, then one line per
 * change — `~ [blockId] modified: …`, `+ [blockId] added: …`,
 * `- [blockId] deleted: …`. A modified block's segments render inline using
 * `git diff --word-diff` porcelain conventions (`[-removed-]`, `{+added+}`,
 * plain text for unchanged runs, truncated to ~30 characters on each side
 * when long). The format is stable across calls — do not change it without
 * checking `compare.test.ts`.
 */
export const formatVersionDiffForLLM = (diff: FolioAgentVersionDiff): string => {
  const { added, deleted, modified, unchanged } = diff.summaryCounts;
  const header = `Version diff: ${added} added, ${deleted} deleted, ${modified} modified, ${unchanged} unchanged`;
  return [header, ...diff.changes.map(formatChangeLine)].join("\n");
};
