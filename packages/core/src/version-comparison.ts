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
 * Compatible body and table/container segments are aligned by the neutral
 * comparison core. Within each compatible segment, blocks are paired in
 * three passes, each only considering blocks the previous pass left unpaired:
 *
 * 1. **Stable-id pairing.** Blocks whose ids are equal and whose snapshot
 *    provenance marks those ids stable are paired directly. A source
 *    `w14:paraId` is stable identity independent of text, so an equal-id pair
 *    with different text is a genuine edit (`modified`). Deterministically
 *    synthesized ids and `seq-NNNN` fallbacks are positional instead: the
 *    adapter preserves that provenance so an ordinal shift falls through to
 *    exact-text alignment rather than becoming a false identity match.
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
 * design). The neutral comparison core re-classifies eligible exact and
 * closely edited pairs as `movedFrom` / `movedTo` entries sharing a
 * `moveGroupId`. Candidate counts and similarity work share bounded budgets
 * across every story in the package comparison.
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
import type { FolioAIBlock } from "./ai-edits/types";
import type { WordDiffSegment } from "./ai-edits/word-diff";
import {
  compareAlignedFolioContent,
  createContentComparisonWorkSession,
  type FolioContentComparisonWorkSession,
  type FolioContentFormattingChange,
} from "./compare/content";
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
  alignFolioContentStructure,
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

const folioAIBlockIdStability = ({
  id,
  idStability,
}: FolioAIBlock): "stable" | "positional" =>
  idStability ?? (getFolioParaIdFromBlockId(id) === null ? "positional" : "stable");

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
    stableIdMismatch: "pair",
    idStability: folioAIBlockIdStability,
  });
  lcsBudget.remainingCells = workSession.remainingLcsCells;
  return events;
};

const legacySegments = (
  segments: readonly WordDiffSegment[],
): FolioVersionDiffSegment[] => segments.map(({ type, text }) => ({ type, text }));

const legacyFormattingProperties = (
  formatting: FolioContentFormattingChange,
): FolioFormatProperty[] =>
  FORMAT_PROPERTIES.filter((property) =>
    formatting.ranges.some(({ formatting: range }) => range[property] !== undefined),
  );

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
  workSession: FolioContentComparisonWorkSession;
};

const compareStoryBlocks = ({
  baseStory,
  revisedStory,
  baseBlocks,
  revisedBlocks,
  firstMoveGroupId,
  includeText,
  includeFormatting,
  workSession,
}: CompareStoryBlocksOptions): FolioStoryDiff => {
  const steps = alignFolioContentStructure({
    baseBlocks,
    revisedBlocks,
    workSession: workSession.alignment,
    stableIdMismatch: "pair",
    idStability: folioAIBlockIdStability,
  });
  const compared = compareAlignedFolioContent({
    baseBlocks,
    revisedBlocks,
    steps,
    workSession,
    idStability: folioAIBlockIdStability,
    // The legacy version-diff surface has no result-size error in its contract.
    maxChanges: Number.MAX_SAFE_INTEGER,
  });
  if (compared.isErr()) {
    return panic("A version comparison exceeded an unreachable internal result limit", {
      cause: compared.error,
    });
  }

  const changes: FolioBlockDiff[] = [];
  const counts = createSummaryCounts();
  const baseHandle = (block: FolioAIBlock): FolioVersionBlockHandle => {
    if (!baseStory) {
      return panic("A neutral comparison event requires a base story handle");
    }
    return { story: baseStory, blockId: block.id };
  };
  const revisedHandle = (block: FolioAIBlock): FolioVersionBlockHandle => {
    if (!revisedStory) {
      return panic("A neutral comparison event requires a revised story handle");
    }
    return { story: revisedStory, blockId: block.id };
  };
  const addModified = ({
    baseBlock,
    revisedBlock,
    segments,
  }: {
    baseBlock: FolioAIBlock;
    revisedBlock: FolioAIBlock;
    segments: readonly WordDiffSegment[];
  }): void => {
    counts.modified++;
    changes.push({
      type: "modified",
      blockId: revisedBlock.id,
      kind: revisedBlock.kind,
      segments: legacySegments(segments),
      baseHandle: baseHandle(baseBlock),
      revisedHandle: revisedHandle(revisedBlock),
    });
  };
  const addFormattingOrUnchanged = ({
    baseBlock,
    revisedBlock,
    formatting,
  }: {
    baseBlock: FolioAIBlock;
    revisedBlock: FolioAIBlock;
    formatting: FolioContentFormattingChange | undefined;
  }): void => {
    const changedProperties =
      includeFormatting && formatting ? legacyFormattingProperties(formatting) : [];
    if (changedProperties.length === 0) {
      counts.unchanged++;
      return;
    }
    counts.formatChanged++;
    changes.push({
      type: "formatChanged",
      blockId: revisedBlock.id,
      kind: revisedBlock.kind,
      text: revisedBlock.text,
      changedProperties,
      baseHandle: baseHandle(baseBlock),
      revisedHandle: revisedHandle(revisedBlock),
    });
  };
  const addDeleted = (block: FolioAIBlock): void => {
    counts.deleted++;
    changes.push({
      type: "deleted",
      blockId: block.id,
      kind: block.kind,
      text: block.text,
      baseHandle: baseHandle(block),
    });
  };
  const addInserted = (block: FolioAIBlock): void => {
    counts.added++;
    changes.push({
      type: "added",
      blockId: block.id,
      kind: block.kind,
      text: block.text,
      revisedHandle: revisedHandle(block),
    });
  };

  for (const event of compared.value.events) {
    switch (event.type) {
      case "unchanged":
        counts.unchanged++;
        break;
      case "formatting":
        addFormattingOrUnchanged({
          baseBlock: event.baseBlocks[0],
          revisedBlock: event.revisedBlocks[0],
          formatting: event.formatting,
        });
        break;
      case "modified": {
        const baseBlock = event.baseBlocks[0];
        const revisedBlock = event.revisedBlocks[0];
        if (baseBlock.text !== revisedBlock.text) {
          if (includeText) {
            addModified({ baseBlock, revisedBlock, segments: event.segments });
          } else {
            counts.unchanged++;
          }
          break;
        }
        addFormattingOrUnchanged({ baseBlock, revisedBlock, formatting: event.formatting });
        break;
      }
      case "deleted":
        if (includeText) {
          addDeleted(event.baseBlocks[0]);
        }
        break;
      case "inserted":
        if (includeText) {
          addInserted(event.revisedBlocks[0]);
        }
        break;
      case "movedFrom":
        if (includeText) {
          const block = event.baseBlocks[0];
          changes.push({
            type: "movedFrom",
            blockId: block.id,
            kind: block.kind,
            text: block.text,
            moveGroupId: firstMoveGroupId + event.moveId - 1,
            baseHandle: baseHandle(block),
          });
        }
        break;
      case "movedTo":
        if (includeText) {
          const block = event.revisedBlocks[0];
          changes.push({
            type: "movedTo",
            blockId: block.id,
            kind: block.kind,
            text: block.text,
            moveGroupId: firstMoveGroupId + event.moveId - 1,
            revisedHandle: revisedHandle(block),
          });
          counts.moved++;
        }
        break;
      case "split": {
        if (!includeText) {
          counts.unchanged++;
          break;
        }
        const baseBlock = event.baseBlocks[0];
        const [firstRevised, secondRevised] = event.revisedBlocks;
        addModified({
          baseBlock,
          revisedBlock: firstRevised,
          segments: workSession.diffText(baseBlock.text, firstRevised.text),
        });
        addInserted(secondRevised);
        break;
      }
      case "merge": {
        if (!includeText) {
          counts.unchanged++;
          break;
        }
        const [firstBase, secondBase] = event.baseBlocks;
        const revisedBlock = event.revisedBlocks[0];
        addModified({
          baseBlock: firstBase,
          revisedBlock,
          segments: workSession.diffText(firstBase.text, revisedBlock.text),
        });
        addDeleted(secondBase);
        break;
      }
      default: {
        const unreachable: never = event;
        panic("Unhandled neutral content comparison event", { event: unreachable });
      }
    }
  }

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
  // One neutral comparison session owns the aggregate alignment, word-diff,
  // and move allowances across every story in this package comparison.
  const comparisonWorkSession = createContentComparisonWorkSession();

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
      workSession: comparisonWorkSession,
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
