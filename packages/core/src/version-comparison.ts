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
 *
 * ## Move detection
 *
 * Relocated content would otherwise report as an unrelated `deleted` +
 * `added` pair (both order-preserving passes drop crossing matches by
 * design). A post-pass re-classifies such pairs: an `added` and a `deleted`
 * block with identical text and at least {@link MOVE_MINIMUM_WORD_COUNT}
 * words become `movedFrom` / `movedTo` entries sharing a `moveGroupId`. The
 * word-count floor keeps boilerplate one-liners ("Confidential", empty
 * headings) from pairing as spurious moves. Blocks a positional zip already
 * mis-paired as `modified` are out of this pass's reach — a known limitation
 * of the gap fallback, not of the move pass.
 *
 * ## Format-only changes
 *
 * A paired block whose text is byte-equal but whose run-level formatting
 * (bold, italic, underline, strike, font family, font size, color) differs
 * reports as `formatChanged` with the set of properties that differ, instead
 * of silently counting as `unchanged`. Detection walks the two blocks'
 * preview runs character-aligned; when a block carries non-text inline
 * content that makes the preview texts disagree, detection backs off to
 * `unchanged` rather than misattribute properties.
 */

import { FolioDocxReviewer } from "./ai-edits/headless";
import type { FolioAIBlock, FolioAIBlockPreviewRun } from "./ai-edits/types";
import { diffWordSegments, type WordDiffSegment } from "./ai-edits/word-diff";
import { getFolioParaIdFromBlockId } from "./types/block-id";

/** One word-level diff segment within a `modified` block. Mirrors {@link WordDiffSegment}. */
export type FolioVersionDiffSegment = WordDiffSegment;

/** Run-level formatting properties compared for `formatChanged` detection. */
const FORMAT_PROPERTIES = [
  "bold",
  "italic",
  "underline",
  "strike",
  "fontFamily",
  "fontSizePt",
  "color",
] as const;

/** A run-level formatting property that can differ in a `formatChanged` block. */
export type FolioFormatProperty = (typeof FORMAT_PROPERTIES)[number];

/** One block-level change between two document versions, in revised-side document order. */
export type FolioBlockDiff =
  | { type: "added"; blockId: string; kind: string; text: string }
  | { type: "deleted"; blockId: string; kind: string; text: string }
  | { type: "modified"; blockId: string; kind: string; segments: FolioVersionDiffSegment[] }
  | {
      type: "formatChanged";
      blockId: string;
      kind: string;
      text: string;
      changedProperties: FolioFormatProperty[];
    }
  | { type: "movedFrom"; blockId: string; kind: string; text: string; moveGroupId: number }
  | { type: "movedTo"; blockId: string; kind: string; text: string; moveGroupId: number };

/** Result of {@link compareDocxVersions}. */
export type FolioVersionDiff = {
  /** Every changed block, in revised-side document order (deletions and move sources slotted where they sat). */
  changes: FolioBlockDiff[];
  /** Counts across every paired/unpaired block, including the unchanged blocks `changes` omits. `moved` counts pairs, not entries. */
  summaryCounts: {
    added: number;
    deleted: number;
    modified: number;
    formatChanged: number;
    moved: number;
    unchanged: number;
  };
};

type BlockPair = { baseIndex: number; revisedIndex: number };
type IndexedBlock = { block: FolioAIBlock; index: number };

const isStableBlockId = (id: string): boolean => getFolioParaIdFromBlockId(id) !== null;

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
  const lengths = new Int32Array(pairs.length).fill(1);
  const predecessors = new Int32Array(pairs.length).fill(-1);
  let bestEnd = 0;
  for (let i = 0; i < pairs.length; i++) {
    for (let j = 0; j < i; j++) {
      const current = pairs[j];
      const candidate = pairs[i];
      if (!current || !candidate) {
        continue;
      }
      if (
        current.revisedIndex < candidate.revisedIndex &&
        (lengths[j] ?? 0) + 1 > (lengths[i] ?? 0)
      ) {
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
const pairByStableId = (
  base: readonly FolioAIBlock[],
  revised: readonly FolioAIBlock[],
): BlockPair[] => {
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
export const exceedsLcsBudget = (
  unpairedBaseCount: number,
  unpairedRevisedCount: number,
): boolean => unpairedBaseCount * unpairedRevisedCount > MAX_LCS_CELLS;

/** Pass 2: order-preserving LCS by exact text equality over the blocks pass 1 left unpaired. */
const pairByExactText = (
  base: readonly IndexedBlock[],
  revised: readonly IndexedBlock[],
): BlockPair[] => {
  const m = base.length;
  const n = revised.length;
  if (m === 0 || n === 0) {
    return [];
  }
  if (exceedsLcsBudget(m, n)) {
    return [];
  }
  const baseTexts = base.map(({ block }) => block.text);
  const revisedTexts = revised.map(({ block }) => block.text);
  // A single flat Int32Array (indexed `i * (n + 1) + j`) instead of `m + 1`
  // separately-allocated rows: one contiguous allocation instead of thousands
  // of small ones, which matters once `m` and `n` approach the budget above.
  const stride = n + 1;
  const dp = new Int32Array((m + 1) * stride);
  for (let i = m - 1; i >= 0; i--) {
    const rowOffset = i * stride;
    const nextRowOffset = (i + 1) * stride;
    const baseText = baseTexts[i];
    for (let j = n - 1; j >= 0; j--) {
      const revisedText = revisedTexts[j];
      dp[rowOffset + j] =
        baseText === revisedText
          ? (dp[nextRowOffset + j + 1] ?? 0) + 1
          : Math.max(dp[nextRowOffset + j] ?? 0, dp[rowOffset + j + 1] ?? 0);
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
    if (baseTexts[i] === revisedTexts[j]) {
      pairs.push({ baseIndex: baseEntry.index, revisedIndex: revisedEntry.index });
      i++;
      j++;
      continue;
    }
    const down = dp[(i + 1) * stride + j] ?? 0;
    const right = dp[i * stride + j + 1] ?? 0;
    if (down >= right) {
      i++;
    } else {
      j++;
    }
  }
  return pairs;
};

/**
 * One step of a completed alignment, in revised-side document order with
 * base-only blocks slotted where they sat. `pair` events cover pass 1/2
 * anchors and pass 3's positional zip alike; whether the pair is unchanged,
 * modified, or format-changed is the consumer's call.
 */
export type FolioAlignedBlockEvent =
  | { type: "pair"; baseBlock: FolioAIBlock; revisedBlock: FolioAIBlock }
  | { type: "baseOnly"; block: FolioAIBlock }
  | { type: "revisedOnly"; block: FolioAIBlock };

/**
 * Run the three-pass alignment (see the module doc comment) over two block
 * snapshots and flatten it into an ordered event stream. Shared by
 * {@link compareDocxVersions} and the redline generator so both interpret
 * one document walk instead of re-deriving it.
 */
export const alignFolioBlocks = (
  baseBlocks: readonly FolioAIBlock[],
  revisedBlocks: readonly FolioAIBlock[],
): FolioAlignedBlockEvent[] => {
  const stableIdAnchors = pairByStableId(baseBlocks, revisedBlocks);
  const usedBaseIndexes = new Set(stableIdAnchors.map((anchor) => anchor.baseIndex));
  const usedRevisedIndexes = new Set(stableIdAnchors.map((anchor) => anchor.revisedIndex));

  const baseRemaining: IndexedBlock[] = [];
  baseBlocks.forEach((block, blockIndex) => {
    if (!usedBaseIndexes.has(blockIndex)) {
      baseRemaining.push({ block, index: blockIndex });
    }
  });
  const revisedRemaining: IndexedBlock[] = [];
  revisedBlocks.forEach((block, blockIndex) => {
    if (!usedRevisedIndexes.has(blockIndex)) {
      revisedRemaining.push({ block, index: blockIndex });
    }
  });
  const exactTextAnchors = pairByExactText(baseRemaining, revisedRemaining);

  const anchors = longestIncreasingByRevisedIndex(
    [...stableIdAnchors, ...exactTextAnchors].toSorted((a, b) => a.baseIndex - b.baseIndex),
  );

  const events: FolioAlignedBlockEvent[] = [];

  /** Pass 3: positionally zip the leftover blocks in one gap between anchors. */
  const emitGap = (
    baseFrom: number,
    baseTo: number,
    revisedFrom: number,
    revisedTo: number,
  ): void => {
    const pairedCount = Math.min(baseTo - baseFrom, revisedTo - revisedFrom);
    for (let k = 0; k < pairedCount; k++) {
      const baseBlock = baseBlocks[baseFrom + k];
      const revisedBlock = revisedBlocks[revisedFrom + k];
      if (baseBlock && revisedBlock) {
        events.push({ type: "pair", baseBlock, revisedBlock });
      }
    }
    for (let k = baseFrom + pairedCount; k < baseTo; k++) {
      const block = baseBlocks[k];
      if (block) {
        events.push({ type: "baseOnly", block });
      }
    }
    for (let k = revisedFrom + pairedCount; k < revisedTo; k++) {
      const block = revisedBlocks[k];
      if (block) {
        events.push({ type: "revisedOnly", block });
      }
    }
  };

  let baseCursor = 0;
  let revisedCursor = 0;
  for (const anchor of anchors) {
    emitGap(baseCursor, anchor.baseIndex, revisedCursor, anchor.revisedIndex);
    const baseBlock = baseBlocks[anchor.baseIndex];
    const revisedBlock = revisedBlocks[anchor.revisedIndex];
    if (baseBlock && revisedBlock) {
      events.push({ type: "pair", baseBlock, revisedBlock });
    }
    baseCursor = anchor.baseIndex + 1;
    revisedCursor = anchor.revisedIndex + 1;
  }
  emitGap(baseCursor, baseBlocks.length, revisedCursor, revisedBlocks.length);

  return events;
};

const previewRunsText = (runs: readonly FolioAIBlockPreviewRun[]): string =>
  runs.map((run) => run.text).join("");

/**
 * Character-aligned formatting diff of two text-equal blocks. Walks both
 * blocks' preview runs in parallel and collects every property whose value
 * differs anywhere in the overlap. A side with `previewRuns === undefined`
 * (the snapshot omits them when every run is unstyled) counts as one
 * unstyled run spanning the whole text. When both sides carry runs but
 * their concatenated texts disagree (non-text inline content can make the
 * preview text drift from the block text), positions can't be aligned, so
 * detection backs off and reports no change.
 */
const diffPreviewRunFormatting = (
  base: FolioAIBlock,
  revised: FolioAIBlock,
): FolioFormatProperty[] => {
  if (base.previewRuns === undefined && revised.previewRuns === undefined) {
    return [];
  }
  const baseText = base.previewRuns === undefined ? null : previewRunsText(base.previewRuns);
  const revisedText =
    revised.previewRuns === undefined ? null : previewRunsText(revised.previewRuns);
  if (baseText !== null && revisedText !== null && baseText !== revisedText) {
    return [];
  }
  const text = baseText ?? revisedText;
  if (text === null || text.length === 0) {
    return [];
  }
  const baseRuns: readonly FolioAIBlockPreviewRun[] = base.previewRuns ?? [{ text }];
  const revisedRuns: readonly FolioAIBlockPreviewRun[] = revised.previewRuns ?? [{ text }];

  const changed = new Set<FolioFormatProperty>();
  let baseRunIndex = 0;
  let revisedRunIndex = 0;
  let baseOffset = 0;
  let revisedOffset = 0;
  while (baseRunIndex < baseRuns.length && revisedRunIndex < revisedRuns.length) {
    const baseRun = baseRuns[baseRunIndex];
    const revisedRun = revisedRuns[revisedRunIndex];
    if (!baseRun || !revisedRun) {
      break;
    }
    const step = Math.min(baseRun.text.length - baseOffset, revisedRun.text.length - revisedOffset);
    if (step > 0) {
      for (const property of FORMAT_PROPERTIES) {
        if (baseRun[property] !== revisedRun[property]) {
          changed.add(property);
        }
      }
    }
    baseOffset += step;
    revisedOffset += step;
    if (baseOffset >= baseRun.text.length) {
      baseRunIndex++;
      baseOffset = 0;
    }
    if (revisedOffset >= revisedRun.text.length) {
      revisedRunIndex++;
      revisedOffset = 0;
    }
  }
  // Report in the stable FORMAT_PROPERTIES order, not set-insertion order.
  return FORMAT_PROPERTIES.filter((property) => changed.has(property));
};

/**
 * Floor for move detection: an added/deleted text must hold at least this
 * many whitespace-separated words before an identical pair re-classifies as
 * a move. Short boilerplate ("Confidential", a bare heading word) recurs
 * throughout real documents and would otherwise pair as spurious moves.
 */
const MOVE_MINIMUM_WORD_COUNT = 3;

const meetsMoveWordCount = (text: string): boolean => {
  // Iterate matches instead of split(): a large block would otherwise
  // allocate its entire token array just to count to the floor.
  const words = text.matchAll(/\S+/gu);
  let count = 0;
  while (!words.next().done) {
    if (++count >= MOVE_MINIMUM_WORD_COUNT) {
      return true;
    }
  }
  return false;
};

/**
 * Re-classify `deleted` + `added` pairs with identical text as
 * `movedFrom` / `movedTo` entries sharing a `moveGroupId`, in place, so each
 * side keeps its slot in the revised-side document order. Matching is FIFO
 * per text, so duplicated boilerplate above the word floor pairs
 * first-to-first rather than fanning out.
 */
const detectMoves = (
  changes: FolioBlockDiff[],
  counts: FolioVersionDiff["summaryCounts"],
): void => {
  const deletedIndexesByText = new Map<string, number[]>();
  changes.forEach((change, index) => {
    if (change.type === "deleted" && meetsMoveWordCount(change.text)) {
      const queue = deletedIndexesByText.get(change.text) ?? [];
      queue.push(index);
      deletedIndexesByText.set(change.text, queue);
    }
  });
  if (deletedIndexesByText.size === 0) {
    return;
  }

  let moveGroupId = 0;
  changes.forEach((change, index) => {
    if (change.type !== "added") {
      return;
    }
    const deletedIndex = deletedIndexesByText.get(change.text)?.shift();
    if (deletedIndex === undefined) {
      return;
    }
    const deleted = changes[deletedIndex];
    if (!deleted || deleted.type !== "deleted") {
      return;
    }
    moveGroupId++;
    changes[deletedIndex] = {
      type: "movedFrom",
      blockId: deleted.blockId,
      kind: deleted.kind,
      text: deleted.text,
      moveGroupId,
    };
    changes[index] = {
      type: "movedTo",
      blockId: change.blockId,
      kind: change.kind,
      text: change.text,
      moveGroupId,
    };
    counts.deleted--;
    counts.added--;
    counts.moved++;
  });
};

/**
 * Compare two `.docx` buffers and return a structured, block-level diff.
 * See the module doc comment for the as-accepted comparison semantics, the
 * three-pass alignment algorithm, move detection, and format-only change
 * detection.
 */
export const compareDocxVersions = async (
  base: ArrayBuffer,
  revised: ArrayBuffer,
): Promise<FolioVersionDiff> => {
  const [baseReviewer, revisedReviewer] = await Promise.all([
    FolioDocxReviewer.fromBuffer(base),
    FolioDocxReviewer.fromBuffer(revised),
  ]);
  const baseBlocks = baseReviewer.snapshot().blocks;
  const revisedBlocks = revisedReviewer.snapshot().blocks;

  const changes: FolioBlockDiff[] = [];
  const counts: FolioVersionDiff["summaryCounts"] = {
    added: 0,
    deleted: 0,
    modified: 0,
    formatChanged: 0,
    moved: 0,
    unchanged: 0,
  };

  for (const event of alignFolioBlocks(baseBlocks, revisedBlocks)) {
    if (event.type === "pair") {
      const { baseBlock, revisedBlock } = event;
      if (baseBlock.text !== revisedBlock.text) {
        counts.modified++;
        changes.push({
          type: "modified",
          blockId: revisedBlock.id,
          kind: revisedBlock.kind,
          segments: diffWordSegments(baseBlock.text, revisedBlock.text),
        });
        continue;
      }
      const changedProperties = diffPreviewRunFormatting(baseBlock, revisedBlock);
      if (changedProperties.length > 0) {
        counts.formatChanged++;
        changes.push({
          type: "formatChanged",
          blockId: revisedBlock.id,
          kind: revisedBlock.kind,
          text: revisedBlock.text,
          changedProperties,
        });
        continue;
      }
      counts.unchanged++;
      continue;
    }
    if (event.type === "baseOnly") {
      counts.deleted++;
      changes.push({
        type: "deleted",
        blockId: event.block.id,
        kind: event.block.kind,
        text: event.block.text,
      });
      continue;
    }
    counts.added++;
    changes.push({
      type: "added",
      blockId: event.block.id,
      kind: event.block.kind,
      text: event.block.text,
    });
  }

  detectMoves(changes, counts);

  return { changes, summaryCounts: counts };
};
