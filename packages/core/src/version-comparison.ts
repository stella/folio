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

import { panic, TaggedError } from "better-result";

import { FolioDocxReviewer, type FolioDocumentStoryHandle } from "./ai-edits/headless";
import type { FolioAIBlock, FolioAIBlockPreviewRun } from "./ai-edits/types";
import { createWordDiffSession, type WordDiffSegment } from "./ai-edits/word-diff";
import { pairFolioDocumentStories, type FolioDocumentStoryPair } from "./document-stories";
import {
  FOLIO_DOCUMENT_METADATA_PROPERTIES,
  FOLIO_DOCUMENT_PRIVACY_TRANSFORMS,
  isFolioDocumentPrivacyTransform,
  PRIVATE_METADATA_PROPERTIES_BY_TRANSFORM,
  type FolioDocumentMetadataProperty,
  type FolioDocumentPrivacyOptions,
  type FolioDocumentPrivacyReport,
  type FolioDocumentPrivacyTransform,
} from "./docx/metadataPrivacy";
import { getFolioParaIdFromBlockId } from "./types/block-id";
import {
  alignFolioContentBlocks,
  createFolioContentAlignmentWorkSession,
  exceedsFolioContentLcsBudget,
  type FolioContentAlignedBlockEvent,
} from "./compare/content-alignment";

/** One word-level diff segment within a `modified` block. Mirrors {@link WordDiffSegment}. */
export type FolioVersionDiffSegment = WordDiffSegment;

/** Independently selectable comparison scopes. */
export const FOLIO_VERSION_COMPARISON_SCOPES = Object.freeze([
  "text",
  "formatting",
  "metadata",
] as const);

export type FolioVersionComparisonScope = (typeof FOLIO_VERSION_COMPARISON_SCOPES)[number];

export const isFolioVersionComparisonScope = (
  value: unknown,
): value is FolioVersionComparisonScope =>
  FOLIO_VERSION_COMPARISON_SCOPES.some((scope) => scope === value);

export type FolioCompareDocxVersionsOptions = {
  /** Selected scopes; defaults to text and formatting. */
  include?: readonly FolioVersionComparisonScope[];
  /** Optional output-only privacy transforms. Source buffers are never mutated. */
  privacy?: FolioVersionDiffPrivacyOptions;
};

export class InvalidFolioVersionComparisonOptionsError extends TaggedError(
  "InvalidFolioVersionComparisonOptionsError",
)<{
  message: string;
  option: "include" | "privacy.transforms";
  receivedValue: unknown;
}> {}

export { FOLIO_DOCUMENT_METADATA_PROPERTIES };
export type { FolioDocumentMetadataProperty };
export type FolioDocumentMetadataValue = string | number | null;

export type FolioMetadataDiff = {
  property: FolioDocumentMetadataProperty;
  baseValue: FolioDocumentMetadataValue;
  revisedValue: FolioDocumentMetadataValue;
};

export const FOLIO_VERSION_COMPARISON_PRIVACY_TRANSFORMS = FOLIO_DOCUMENT_PRIVACY_TRANSFORMS;

export type FolioVersionComparisonPrivacyTransform = FolioDocumentPrivacyTransform;

export const isFolioVersionComparisonPrivacyTransform = (
  value: unknown,
): value is FolioVersionComparisonPrivacyTransform => isFolioDocumentPrivacyTransform(value);

export type FolioVersionDiffPrivacyOptions = FolioDocumentPrivacyOptions;

export type FolioVersionDiffPrivacyReport = FolioDocumentPrivacyReport;

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

/** Stable location of one compared block within its source document. */
export type FolioVersionBlockHandle = {
  story: FolioDocumentStoryHandle;
  blockId: string;
};

/** One block-level change between two document versions, in revised-side document order. */
export type FolioBlockDiff =
  | {
      type: "added";
      blockId: string;
      kind: string;
      text: string;
      revisedHandle: FolioVersionBlockHandle;
    }
  | {
      type: "deleted";
      blockId: string;
      kind: string;
      text: string;
      baseHandle: FolioVersionBlockHandle;
    }
  | {
      type: "modified";
      blockId: string;
      kind: string;
      segments: FolioVersionDiffSegment[];
      baseHandle: FolioVersionBlockHandle;
      revisedHandle: FolioVersionBlockHandle;
    }
  | {
      type: "formatChanged";
      blockId: string;
      kind: string;
      text: string;
      changedProperties: FolioFormatProperty[];
      baseHandle: FolioVersionBlockHandle;
      revisedHandle: FolioVersionBlockHandle;
    }
  | {
      type: "movedFrom";
      blockId: string;
      kind: string;
      text: string;
      moveGroupId: number;
      baseHandle: FolioVersionBlockHandle;
    }
  | {
      type: "movedTo";
      blockId: string;
      kind: string;
      text: string;
      moveGroupId: number;
      revisedHandle: FolioVersionBlockHandle;
    };

export type FolioVersionDiffSummaryCounts = {
  added: number;
  deleted: number;
  modified: number;
  formatChanged: number;
  moved: number;
  metadataChanged: number;
  unchanged: number;
};

/** Changes within one matched, added, or deleted document story. */
export type FolioStoryDiff = {
  baseStory: FolioDocumentStoryHandle | null;
  revisedStory: FolioDocumentStoryHandle | null;
  changes: FolioBlockDiff[];
  summaryCounts: FolioVersionDiffSummaryCounts;
};

/** Result of {@link compareDocxVersions}. */
export type FolioVersionDiff = {
  /** Every changed block, in revised-side document order (deletions and move sources slotted where they sat). */
  changes: FolioBlockDiff[];
  /** Per-story results in base order followed by stories added in the revised document. */
  stories: FolioStoryDiff[];
  /** Changed package metadata fields in stable property order. */
  metadataChanges: FolioMetadataDiff[];
  /** Applied privacy policy and the fields it removed from this result. */
  privacyReport: FolioVersionDiffPrivacyReport;
  /** Counts across every paired/unpaired block, including the unchanged blocks `changes` omits. `moved` counts pairs, not entries. */
  summaryCounts: FolioVersionDiffSummaryCounts;
};

export const exceedsLcsBudget = exceedsFolioContentLcsBudget;

export type FolioVersionComparisonLcsBudget = { remainingCells: number };

const createLcsBudget = (): FolioVersionComparisonLcsBudget => {
  const session = createFolioContentAlignmentWorkSession();
  return { remainingCells: session.remainingLcsCells };
};

export type FolioAlignedBlockEvent = FolioContentAlignedBlockEvent<FolioAIBlock>;

/**
 * Compatibility adapter for the DOCX snapshot comparison surface.
 *
 * Older snapshots encode id stability in the id shape, so the adapter resolves
 * that policy while the representation-neutral core can treat caller ids as
 * stable by default.
 */
export const alignFolioBlocks = (
  baseBlocks: readonly FolioAIBlock[],
  revisedBlocks: readonly FolioAIBlock[],
  lcsBudget: FolioVersionComparisonLcsBudget = createLcsBudget(),
): FolioAlignedBlockEvent[] => {
  const workSession = { remainingLcsCells: lcsBudget.remainingCells };
  const events = alignFolioContentBlocks(baseBlocks, revisedBlocks, {
    workSession,
    idStability: ({ id }) => (getFolioParaIdFromBlockId(id) === null ? "positional" : "stable"),
  });
  lcsBudget.remainingCells = workSession.remainingLcsCells;
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
 * Cap on how many same-text `deleted` candidates {@link detectMoves} queues
 * per distinct text. Duplicated boilerplate at or above the word floor
 * (e.g. a repeated long clause) could otherwise grow one text's candidate
 * list without bound; past this cap, further same-text deletions are simply
 * left as `deleted` (never matched to a move) instead of queued.
 */
const MAX_MOVE_CANDIDATES_PER_TEXT = 10_000;

/**
 * FIFO queue of `changes` indexes for one deleted text. `shift()` on a plain
 * array is O(k); a head cursor makes dequeue O(1) so `detectMoves` stays
 * linear in the number of added/deleted blocks instead of quadratic when a
 * document repeats the same text many times.
 */
type DeletedIndexQueue = { items: number[]; head: number };

const dequeueDeletedIndex = (queue: DeletedIndexQueue | undefined): number | undefined => {
  if (!queue || queue.head >= queue.items.length) {
    return undefined;
  }
  const index = queue.items[queue.head];
  queue.head += 1;
  return index;
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
  counts: FolioVersionDiffSummaryCounts,
  firstMoveGroupId: number,
): void => {
  const deletedIndexesByText = new Map<string, DeletedIndexQueue>();
  changes.forEach((change, index) => {
    if (change.type !== "deleted" || !meetsMoveWordCount(change.text)) {
      return;
    }
    const queue = deletedIndexesByText.get(change.text);
    if (!queue) {
      deletedIndexesByText.set(change.text, { items: [index], head: 0 });
      return;
    }
    // This build pass runs to completion before any dequeue below, so `head`
    // is always 0 here; comparing against the cap directly is equivalent to
    // (and simpler than) tracking the unconsumed remainder.
    if (queue.items.length < MAX_MOVE_CANDIDATES_PER_TEXT) {
      queue.items.push(index);
    }
  });
  if (deletedIndexesByText.size === 0) {
    return;
  }

  let moveGroupId = firstMoveGroupId - 1;
  changes.forEach((change, index) => {
    if (change.type !== "added") {
      return;
    }
    const deletedIndex = dequeueDeletedIndex(deletedIndexesByText.get(change.text));
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
      baseHandle: deleted.baseHandle,
    };
    changes[index] = {
      type: "movedTo",
      blockId: change.blockId,
      kind: change.kind,
      text: change.text,
      moveGroupId,
      revisedHandle: change.revisedHandle,
    };
    counts.deleted--;
    counts.added--;
    counts.moved++;
  });
};

const createSummaryCounts = (): FolioVersionDiffSummaryCounts => ({
  added: 0,
  deleted: 0,
  modified: 0,
  formatChanged: 0,
  moved: 0,
  metadataChanged: 0,
  unchanged: 0,
});

const addSummaryCounts = (
  target: FolioVersionDiffSummaryCounts,
  source: FolioVersionDiffSummaryCounts,
): void => {
  target.added += source.added;
  target.deleted += source.deleted;
  target.modified += source.modified;
  target.formatChanged += source.formatChanged;
  target.moved += source.moved;
  target.metadataChanged += source.metadataChanged;
  target.unchanged += source.unchanged;
};

type CompareStoryBlocksOptions = FolioDocumentStoryPair & {
  baseBlocks: readonly FolioAIBlock[];
  revisedBlocks: readonly FolioAIBlock[];
  firstMoveGroupId: number;
  includeText: boolean;
  includeFormatting: boolean;
  lcsBudget: FolioVersionComparisonLcsBudget;
  diffText: ReturnType<typeof createWordDiffSession>["diff"];
};

const compareStoryBlocks = ({
  baseStory,
  revisedStory,
  baseBlocks,
  revisedBlocks,
  firstMoveGroupId,
  includeText,
  includeFormatting,
  lcsBudget,
  diffText,
}: CompareStoryBlocksOptions): FolioStoryDiff => {
  const changes: FolioBlockDiff[] = [];
  const counts = createSummaryCounts();

  for (const event of alignFolioBlocks(baseBlocks, revisedBlocks, lcsBudget)) {
    if (event.type === "pair") {
      if (!baseStory || !revisedStory) {
        panic("A paired comparison event requires both story handles");
      }
      const { baseBlock, revisedBlock } = event;
      const baseHandle = { story: baseStory, blockId: baseBlock.id };
      const revisedHandle = { story: revisedStory, blockId: revisedBlock.id };
      if (baseBlock.text !== revisedBlock.text) {
        if (!includeText) {
          counts.unchanged++;
          continue;
        }
        counts.modified++;
        changes.push({
          type: "modified",
          blockId: revisedBlock.id,
          kind: revisedBlock.kind,
          segments: diffText(baseBlock.text, revisedBlock.text),
          baseHandle,
          revisedHandle,
        });
        continue;
      }
      const changedProperties = includeFormatting
        ? diffPreviewRunFormatting(baseBlock, revisedBlock)
        : [];
      if (changedProperties.length > 0) {
        counts.formatChanged++;
        changes.push({
          type: "formatChanged",
          blockId: revisedBlock.id,
          kind: revisedBlock.kind,
          text: revisedBlock.text,
          changedProperties,
          baseHandle,
          revisedHandle,
        });
        continue;
      }
      counts.unchanged++;
      continue;
    }
    if (event.type === "baseOnly") {
      if (!includeText) {
        continue;
      }
      if (!baseStory) {
        panic("A base-only comparison event requires a base story handle");
      }
      counts.deleted++;
      changes.push({
        type: "deleted",
        blockId: event.block.id,
        kind: event.block.kind,
        text: event.block.text,
        baseHandle: { story: baseStory, blockId: event.block.id },
      });
      continue;
    }
    if (!includeText) {
      continue;
    }
    if (!revisedStory) {
      panic("A revised-only comparison event requires a revised story handle");
    }
    counts.added++;
    changes.push({
      type: "added",
      blockId: event.block.id,
      kind: event.block.kind,
      text: event.block.text,
      revisedHandle: { story: revisedStory, blockId: event.block.id },
    });
  }

  detectMoves(changes, counts, firstMoveGroupId);
  return { baseStory, revisedStory, changes, summaryCounts: counts };
};

const DEFAULT_COMPARISON_SCOPES = Object.freeze([
  "text",
  "formatting",
] as const satisfies readonly FolioVersionComparisonScope[]);

const resolveComparisonScopes = (
  options: FolioCompareDocxVersionsOptions,
): ReadonlySet<FolioVersionComparisonScope> => {
  const include = options.include ?? DEFAULT_COMPARISON_SCOPES;
  if (include.length === 0 || include.some((scope) => !isFolioVersionComparisonScope(scope))) {
    throw new InvalidFolioVersionComparisonOptionsError({
      message: "Version comparison requires at least one recognized scope",
      option: "include",
      receivedValue: include,
    });
  }
  return new Set(include);
};

const resolvePrivacyTransforms = (
  transforms: unknown,
): FolioVersionComparisonPrivacyTransform[] => {
  if (
    !Array.isArray(transforms) ||
    transforms.some((transform) => !isFolioVersionComparisonPrivacyTransform(transform))
  ) {
    throw new InvalidFolioVersionComparisonOptionsError({
      message: "Version comparison received an unrecognized privacy transform",
      option: "privacy.transforms",
      receivedValue: transforms,
    });
  }
  const requested = new Set(transforms);
  return FOLIO_VERSION_COMPARISON_PRIVACY_TRANSFORMS.filter((transform) =>
    requested.has(transform),
  );
};

/** Apply auditable, output-only privacy transforms to a structured version diff. */
export const applyFolioVersionDiffPrivacy = (
  diff: FolioVersionDiff,
  options: FolioVersionDiffPrivacyOptions,
): FolioVersionDiff => {
  const requestedTransforms = resolvePrivacyTransforms(options.transforms);
  const appliedTransformSet = new Set([
    ...diff.privacyReport.appliedTransforms,
    ...requestedTransforms,
  ]);
  const appliedTransforms = FOLIO_VERSION_COMPARISON_PRIVACY_TRANSFORMS.filter((transform) =>
    appliedTransformSet.has(transform),
  );
  const removedPropertySet = new Set<FolioDocumentMetadataProperty>();
  for (const transform of appliedTransforms) {
    for (const property of PRIVATE_METADATA_PROPERTIES_BY_TRANSFORM[transform]) {
      removedPropertySet.add(property);
    }
  }
  const actuallyRemovedPropertySet = new Set([
    ...diff.privacyReport.removedMetadataProperties,
    ...diff.metadataChanges
      .filter(({ property }) => removedPropertySet.has(property))
      .map(({ property }) => property),
  ]);
  const removedMetadataProperties = FOLIO_DOCUMENT_METADATA_PROPERTIES.filter((property) =>
    actuallyRemovedPropertySet.has(property),
  );
  const metadataChanges = diff.metadataChanges.filter(
    ({ property }) => !removedPropertySet.has(property),
  );

  return {
    ...diff,
    metadataChanges,
    privacyReport: { appliedTransforms, removedMetadataProperties },
    summaryCounts: { ...diff.summaryCounts, metadataChanged: metadataChanges.length },
  };
};

type DocumentProperties = ReturnType<FolioDocxReviewer["getDocumentProperties"]>;

const normalizeMetadataValue = (
  properties: DocumentProperties,
  property: FolioDocumentMetadataProperty,
): FolioDocumentMetadataValue => {
  const value = properties?.[property];
  return value instanceof Date ? value.toISOString() : (value ?? null);
};

const compareMetadata = (
  base: DocumentProperties,
  revised: DocumentProperties,
): FolioMetadataDiff[] => {
  const changes: FolioMetadataDiff[] = [];
  for (const property of FOLIO_DOCUMENT_METADATA_PROPERTIES) {
    const baseValue = normalizeMetadataValue(base, property);
    const revisedValue = normalizeMetadataValue(revised, property);
    if (baseValue !== revisedValue) {
      changes.push({ property, baseValue, revisedValue });
    }
  }
  return changes;
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
  options: FolioCompareDocxVersionsOptions = {},
): Promise<FolioVersionDiff> => {
  const scopes = resolveComparisonScopes(options);
  const [baseReviewer, revisedReviewer] = await Promise.all([
    FolioDocxReviewer.fromBuffer(base),
    FolioDocxReviewer.fromBuffer(revised),
  ]);
  const changes: FolioBlockDiff[] = [];
  const stories: FolioStoryDiff[] = [];
  const counts = createSummaryCounts();
  const baseStories = baseReviewer.listStories().map(({ handle }) => handle);
  const revisedStories = revisedReviewer.listStories().map(({ handle }) => handle);
  let nextMoveGroupId = 1;
  // Shared across every story pair below (not one per story) so an
  // attacker-controlled story count (many footnotes/endnotes) can't force a
  // fresh near-MAX_LCS_CELLS allocation per story. See
  // FolioVersionComparisonLcsBudget's doc comment.
  const lcsBudget = createLcsBudget();
  const wordDiffSession = createWordDiffSession();

  for (const pair of pairFolioDocumentStories(baseStories, revisedStories)) {
    const baseBlocks = pair.baseStory
      ? (baseReviewer.readReviewedStory({ story: pair.baseStory, view: "final" })?.snapshot
          .blocks ?? [])
      : [];
    const revisedBlocks = pair.revisedStory
      ? (revisedReviewer.readReviewedStory({ story: pair.revisedStory, view: "final" })?.snapshot
          .blocks ?? [])
      : [];
    const storyDiff = compareStoryBlocks({
      ...pair,
      baseBlocks,
      revisedBlocks,
      firstMoveGroupId: nextMoveGroupId,
      includeText: scopes.has("text"),
      includeFormatting: scopes.has("formatting"),
      lcsBudget,
      diffText: wordDiffSession.diff,
    });
    stories.push(storyDiff);
    for (const change of storyDiff.changes) {
      changes.push(change);
    }
    addSummaryCounts(counts, storyDiff.summaryCounts);
    nextMoveGroupId += storyDiff.summaryCounts.moved;
  }

  const metadataChanges = scopes.has("metadata")
    ? compareMetadata(baseReviewer.getDocumentProperties(), revisedReviewer.getDocumentProperties())
    : [];
  counts.metadataChanged = metadataChanges.length;

  const diff: FolioVersionDiff = {
    changes,
    stories,
    metadataChanges,
    privacyReport: {
      appliedTransforms: [],
      removedMetadataProperties: [],
    },
    summaryCounts: counts,
  };
  return options.privacy ? applyFolioVersionDiffPrivacy(diff, options.privacy) : diff;
};
