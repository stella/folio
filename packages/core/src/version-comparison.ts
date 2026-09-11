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
 *    O(m·n) table is skipped once the unpaired
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
import type { WordDiffSegment } from "./compare/text-diff";
import {
  compareContentStories,
  createContentComparisonWorkSession,
  FolioContentComparisonLimitError,
  type FolioContentBlockChange,
  type FolioContentComparison,
  type FolioContentComparisonLimit,
  type FolioContentFormattingChange,
  type FolioContentPairRelation,
} from "./compare/content";
import type { FolioContentBlock } from "./compare/content-types";
import { docxBlockToContentInput } from "./compare/docx-content-adapter";
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

/** Aggregate content limit exceeded while comparing all stories in two packages. */
export class FolioVersionComparisonLimitError extends TaggedError(
  "FolioVersionComparisonLimitError",
)<{
  message: string;
  input: "base" | "revised" | "result";
  limit: FolioContentComparisonLimit;
  maximum: number;
  actual: number;
  storyIndex: number;
  blockIndex?: number;
  field?: string;
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

/** A presentation property that can differ in a `formatChanged` block. */
export type FolioFormatProperty = string;

/** A non-presentation block property retained on composite content changes. */
export type FolioBlockProperty = string;

/** One content or presentation property attached to a modified or moved block. */
export type FolioVersionChangeProperty = FolioBlockProperty | FolioFormatProperty;

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
      changedProperties?: FolioVersionChangeProperty[];
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
      segments?: FolioVersionDiffSegment[];
      changedProperties?: FolioVersionChangeProperty[];
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

const projectVersionSegments = (segments: readonly WordDiffSegment[]): FolioVersionDiffSegment[] =>
  segments.map(({ type, text }) => ({ type, text }));

const appendVersionSegment = (
  segments: FolioVersionDiffSegment[],
  segment: FolioVersionDiffSegment,
): void => {
  if (segment.text.length === 0) return;
  const previous = segments.at(-1);
  if (previous?.type === segment.type) {
    previous.text += segment.text;
    return;
  }
  segments.push({ ...segment });
};

const projectRelationSideText = (
  relation: FolioContentPairRelation,
  side: "base" | "revised",
): string =>
  relation.segments
    .filter(({ type }) => type !== (side === "base" ? "ins" : "del"))
    .map(({ text }) => text)
    .join("");

const projectSplitModifiedSegments = (
  first: FolioContentPairRelation,
  second: FolioContentPairRelation,
  separator: FolioContentPairRelation,
): FolioVersionDiffSegment[] => {
  const segments: FolioVersionDiffSegment[] = [];
  for (const segment of projectVersionSegments(first.segments)) {
    appendVersionSegment(segments, segment);
  }
  appendVersionSegment(segments, {
    type: "del",
    text: projectRelationSideText(separator, "base"),
  });
  appendVersionSegment(segments, { type: "del", text: projectRelationSideText(second, "base") });
  return segments;
};

const projectMergeModifiedSegments = (
  first: FolioContentPairRelation,
  second: FolioContentPairRelation,
  separator: FolioContentPairRelation,
): FolioVersionDiffSegment[] => {
  const segments: FolioVersionDiffSegment[] = [];
  for (const segment of projectVersionSegments(first.segments)) {
    appendVersionSegment(segments, segment);
  }
  appendVersionSegment(segments, {
    type: "ins",
    text: projectRelationSideText(separator, "revised"),
  });
  appendVersionSegment(segments, {
    type: "ins",
    text: projectRelationSideText(second, "revised"),
  });
  return segments;
};

const projectVersionFormattingProperties = (
  formatting: FolioContentFormattingChange,
): FolioFormatProperty[] => {
  const properties = new Set<FolioFormatProperty>();
  for (const { key } of formatting.paragraph.authored) properties.add(key);
  for (const { key } of formatting.paragraph.effective) properties.add(key);
  for (const { formatting: range } of formatting.ranges) {
    for (const { key } of range.authored) properties.add(key);
    for (const { key } of range.effective) properties.add(key);
  }
  return [...properties].toSorted();
};

const projectVersionBlockProperties = (
  changes: readonly FolioContentBlockChange[],
): FolioBlockProperty[] => {
  const properties = new Set<FolioBlockProperty>();
  for (const change of changes) {
    switch (change.field) {
      case "blockProperties":
        for (const { key } of change.changes) properties.add(key);
        break;
      case "kind":
      case "structuralBoundaries":
      case "table":
      case "containerPath":
        properties.add(change.field);
        break;
      default: {
        const unreachable: never = change;
        return panic("Unhandled canonical block change", { change: unreachable });
      }
    }
  }
  return [...properties].toSorted();
};

const projectVersionChangeProperties = ({
  blockChanges,
  formatting,
}: {
  blockChanges: readonly FolioContentBlockChange[];
  formatting: FolioContentFormattingChange | undefined;
}): FolioVersionChangeProperty[] => {
  const properties: FolioVersionChangeProperty[] = projectVersionBlockProperties(blockChanges);
  if (formatting) {
    properties.push(...projectVersionFormattingProperties(formatting));
  }
  return properties;
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

type ProjectFolioContentComparisonOptions = FolioDocumentStoryPair & {
  comparison: FolioContentComparison;
  firstMoveGroupId: number;
  includeText: boolean;
  includeFormatting: boolean;
};

const versionComparisonLimitError = (
  cause: FolioContentComparisonLimitError,
  storyIndex: number,
): FolioVersionComparisonLimitError =>
  new FolioVersionComparisonLimitError({
    message: `Document version comparison exceeds the aggregate ${cause.limit} limit.`,
    input: cause.input,
    limit: cause.limit,
    maximum: cause.maximum,
    actual: cause.actual,
    storyIndex,
    ...(cause.blockIndex !== undefined && { blockIndex: cause.blockIndex }),
    ...(cause.field !== undefined && { field: cause.field }),
  });

/** Project the canonical neutral event stream into the structured version-diff surface. @internal */
export const projectFolioContentComparisonToStory = ({
  baseStory,
  revisedStory,
  comparison,
  firstMoveGroupId,
  includeText,
  includeFormatting,
}: ProjectFolioContentComparisonOptions): FolioStoryDiff => {
  const changes: FolioBlockDiff[] = [];
  const counts = createSummaryCounts();
  const baseHandle = (block: FolioContentBlock): FolioVersionBlockHandle => {
    if (!baseStory) {
      return panic("A neutral comparison event requires a base story handle");
    }
    return { story: baseStory, blockId: block.identity.id };
  };
  const revisedHandle = (block: FolioContentBlock): FolioVersionBlockHandle => {
    if (!revisedStory) {
      return panic("A neutral comparison event requires a revised story handle");
    }
    return { story: revisedStory, blockId: block.identity.id };
  };
  const addModified = (
    relation: FolioContentPairRelation,
    segments: readonly WordDiffSegment[] = relation.segments,
  ): void => {
    const baseBlock = relation.base.block;
    const revisedBlock = relation.revised.block;
    const changedProperties = projectVersionChangeProperties({
      blockChanges: relation.blockChanges,
      formatting: includeFormatting ? (relation.formatting ?? undefined) : undefined,
    });
    counts.modified++;
    changes.push({
      type: "modified",
      blockId: revisedBlock.identity.id,
      kind: revisedBlock.kind,
      segments: projectVersionSegments(segments),
      ...(changedProperties.length > 0 && { changedProperties }),
      baseHandle: baseHandle(baseBlock),
      revisedHandle: revisedHandle(revisedBlock),
    });
  };
  const addFormattingOrUnchanged = (relation: FolioContentPairRelation): void => {
    const baseBlock = relation.base.block;
    const revisedBlock = relation.revised.block;
    const changedProperties =
      includeFormatting && relation.formatting
        ? projectVersionFormattingProperties(relation.formatting)
        : [];
    if (changedProperties.length === 0) {
      counts.unchanged++;
      return;
    }
    counts.formatChanged++;
    changes.push({
      type: "formatChanged",
      blockId: revisedBlock.identity.id,
      kind: revisedBlock.kind,
      text: revisedBlock.text,
      changedProperties,
      baseHandle: baseHandle(baseBlock),
      revisedHandle: revisedHandle(revisedBlock),
    });
  };
  const addDeleted = (block: FolioContentBlock): void => {
    counts.deleted++;
    changes.push({
      type: "deleted",
      blockId: block.identity.id,
      kind: block.kind,
      text: block.text,
      baseHandle: baseHandle(block),
    });
  };
  const addInserted = (block: FolioContentBlock): void => {
    counts.added++;
    changes.push({
      type: "added",
      blockId: block.identity.id,
      kind: block.kind,
      text: block.text,
      revisedHandle: revisedHandle(block),
    });
  };

  for (const event of comparison.events) {
    switch (event.type) {
      case "unchanged":
        counts.unchanged++;
        break;
      case "formatting":
        addFormattingOrUnchanged(event.relation);
        break;
      case "modified": {
        if (includeText) {
          addModified(event.relation);
          break;
        }
        addFormattingOrUnchanged(event.relation);
        break;
      }
      case "deleted":
        if (includeText) {
          addDeleted(event.block);
        } else {
          counts.unchanged++;
        }
        break;
      case "inserted":
        if (includeText) {
          addInserted(event.block);
        } else {
          counts.unchanged++;
        }
        break;
      case "movedFrom": {
        if (!includeText) break;
        const block = event.move.relation.base.block;
        changes.push({
          type: "movedFrom",
          blockId: block.identity.id,
          kind: block.kind,
          text: block.text,
          moveGroupId: firstMoveGroupId + event.move.id - 1,
          baseHandle: baseHandle(block),
        });
        break;
      }
      case "movedTo": {
        const relation = event.move.relation;
        const block = relation.revised.block;
        if (!includeText) {
          addFormattingOrUnchanged(relation);
          break;
        }
        const changedProperties = projectVersionChangeProperties({
          blockChanges: relation.blockChanges,
          formatting: includeFormatting ? (relation.formatting ?? undefined) : undefined,
        });
        const hasTextChange = relation.segments.some(({ type }) => type !== "equal");
        changes.push({
          type: "movedTo",
          blockId: block.identity.id,
          kind: block.kind,
          text: block.text,
          moveGroupId: firstMoveGroupId + event.move.id - 1,
          ...(hasTextChange && { segments: projectVersionSegments(relation.segments) }),
          ...(changedProperties.length > 0 && { changedProperties }),
          revisedHandle: revisedHandle(block),
        });
        counts.moved++;
        break;
      }
      case "split": {
        const [first, second] = event.relations;
        if (includeText) {
          addModified(first, projectSplitModifiedSegments(first, second, event.separator));
          // The second paragraph is target content in the combined scope, so
          // its formatting is intrinsic to the added block. Formatting-only
          // projection compares both revised siblings to their shared source.
          addInserted(second.revised.block);
          break;
        }
        addFormattingOrUnchanged(first);
        addFormattingOrUnchanged(second);
        break;
      }
      case "merge": {
        const [first, second] = event.relations;
        if (includeText) {
          addModified(first, projectMergeModifiedSegments(first, second, event.separator));
          addDeleted(second.base.block);
          break;
        }
        addFormattingOrUnchanged(first);
        addFormattingOrUnchanged(second);
        break;
      }
      case "tableReplacement": {
        if (!includeText) {
          counts.unchanged += Math.max(
            event.replacement.baseBlocks.length,
            event.replacement.revisedBlocks.length,
          );
          break;
        }
        for (const block of event.replacement.baseBlocks) addDeleted(block);
        for (const block of event.replacement.revisedBlocks) addInserted(block);
        break;
      }
      case "structural": {
        if (!includeText) {
          counts.unchanged++;
          break;
        }
        const structuralBlock = event.change.blocks.at(event.memberIndex);
        if (!structuralBlock) {
          panic("A structural comparison event has no owned member", {
            type: event.change.type,
            memberIndex: event.memberIndex,
          });
        }
        switch (event.change.type) {
          case "table-delete":
          case "table-row-delete":
          case "table-column-delete":
            addDeleted(structuralBlock);
            break;
          case "table-insert":
          case "table-row-insert":
          case "table-column-insert":
            addInserted(structuralBlock);
            break;
          default: {
            const unreachable: never = event.change;
            panic("Unhandled neutral structural comparison", { change: unreachable });
          }
        }
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
  const pairedStories = pairFolioDocumentStories(baseStories, revisedStories).map(
    (pair) => {
    const baseBlocks = pair.baseStory
      ? (baseReviewer.readReviewedStory({ story: pair.baseStory, view: "final" })?.snapshot
          .blocks.map(docxBlockToContentInput) ?? [])
      : [];
    const revisedBlocks = pair.revisedStory
      ? (revisedReviewer.readReviewedStory({ story: pair.revisedStory, view: "final" })?.snapshot
          .blocks.map(docxBlockToContentInput) ?? [])
      : [];
      return Object.freeze({
        key: Object.freeze({ ...pair }),
        base: Object.freeze({ blocks: baseBlocks }),
        revised: Object.freeze({ blocks: revisedBlocks }),
      });
    },
  );
  const comparedStories = compareContentStories({
    stories: pairedStories,
    workSession: createContentComparisonWorkSession(),
  });
  if (comparedStories.isErr()) {
    const { cause, storyIndex } = comparedStories.error;
    if (cause instanceof FolioContentComparisonLimitError) {
      throw versionComparisonLimitError(cause, storyIndex);
    }
    return panic("A reviewed story produced an invalid content snapshot", { cause, storyIndex });
  }

  for (const { key, comparison } of comparedStories.value) {
    const storyDiff = projectFolioContentComparisonToStory({
      ...key,
      comparison,
      firstMoveGroupId: nextMoveGroupId,
      includeText: scopes.has("text"),
      includeFormatting: scopes.has("formatting"),
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
