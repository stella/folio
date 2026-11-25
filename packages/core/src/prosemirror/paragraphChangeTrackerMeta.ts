/**
 * The transaction meta the paragraph change tracker reads, and the state it
 * keeps. Commands state what an edit did through these setters; only the
 * tracker plugin (`ParagraphChangeTrackerExtension`) reads them.
 */
import type { Node as PMNode } from "prosemirror-model";
import type {
  EditorState,
  Plugin,
  PluginKey,
  Transaction,
} from "prosemirror-state";

import type {
  RemovedSectionReference,
  TrackedSectionEndpointRemoval,
} from "../internal/sectionEndpointResolution";
import type { SectionProperties } from "../types/document";
import { countDocumentStructure } from "./documentStructureCounts";
import { createPluginKeyOwner } from "./pluginKeyOwner";

export const paragraphChangeTrackerKey: Readonly<{
  key: string;
  createKey: () => PluginKey<InternalParagraphChangeTrackerState>;
  owns: (plugin: Plugin<InternalParagraphChangeTrackerState>) => boolean;
  get: (
    state: EditorState,
  ) => Plugin<InternalParagraphChangeTrackerState> | undefined;
  getState: (
    state: EditorState,
  ) => InternalParagraphChangeTrackerState | undefined;
}> = createPluginKeyOwner<InternalParagraphChangeTrackerState>(
  "paragraphChangeTracker",
);

export const CLEAR_META = "clear";
export const IGNORE_META = "ignore";
export const STRUCTURAL_META = "structural";
export const CHANGED_PARAGRAPH_RANGES_META = "folioChangedParagraphRanges";
export const SECTION_ENDPOINT_REMOVAL_META = "folioSectionEndpointRemoval";

export type ChangedParagraphRangeBatch = {
  ranges: readonly { from: number; to: number }[];
  mappingFrom: number;
  replacesStepMapAt?: number;
};

type ChangedParagraphRangesMeta = {
  type: "changed-paragraph-ranges";
  batches: readonly ChangedParagraphRangeBatch[];
};

export const isChangedParagraphRangesMeta = (
  value: unknown,
): value is ChangedParagraphRangesMeta =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  value.type === "changed-paragraph-ranges";

const isRemovedSectionReference = (
  value: unknown,
): value is RemovedSectionReference =>
  typeof value === "object" &&
  value !== null &&
  "part" in value &&
  (value.part === "header" || value.part === "footer") &&
  "type" in value &&
  (value.type === "default" ||
    value.type === "first" ||
    value.type === "even") &&
  "relationshipId" in value &&
  typeof value.relationshipId === "string";

export const isTrackedSectionEndpointRemoval = (
  value: unknown,
): value is TrackedSectionEndpointRemoval =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  value.type === "tracked-section-endpoint-removal" &&
  "sourceParagraphEndpointCount" in value &&
  typeof value.sourceParagraphEndpointCount === "number" &&
  Number.isSafeInteger(value.sourceParagraphEndpointCount) &&
  "expectedParagraphEndpointCount" in value &&
  typeof value.expectedParagraphEndpointCount === "number" &&
  Number.isSafeInteger(value.expectedParagraphEndpointCount) &&
  value.sourceParagraphEndpointCount > value.expectedParagraphEndpointCount &&
  value.expectedParagraphEndpointCount >= 0 &&
  "sourceEndpointFingerprint" in value &&
  typeof value.sourceEndpointFingerprint === "string" &&
  "expectedEndpointFingerprint" in value &&
  typeof value.expectedEndpointFingerprint === "string" &&
  "removedReferences" in value &&
  Array.isArray(value.removedReferences) &&
  value.removedReferences.every(isRemovedSectionReference);

type TrackedSectionEndpointRemovalMeta = {
  type: "tracked-section-endpoint-removal-meta";
  authorization: TrackedSectionEndpointRemoval;
};

export const isTrackedSectionEndpointRemovalMeta = (
  value: unknown,
): value is TrackedSectionEndpointRemovalMeta =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  value.type === "tracked-section-endpoint-removal-meta" &&
  "authorization" in value &&
  isTrackedSectionEndpointRemoval(value.authorization);

export type ParagraphChangeTrackerState = {
  /** Set of paraIds that were modified since last clear */
  changedParaIds: Set<string>;
  /** Whether paragraph order, membership, or block structure changed. */
  structuralChange: boolean;
  /** Whether edited paragraphs still lack IDs or source identity was unavailable. */
  hasUntrackedChanges: boolean;
  /** Paragraph count in the current tracked document. */
  paragraphCount: number;
  /** Cached section-endpoint count used to invalidate stale save authorization. */
  sectionEndpointCount: number;
  /** Exact endpoint owners and properties used to invalidate count-neutral changes. */
  sectionEndpointFingerprint: string;
  /** Exact tracked-resolution transition that may reduce the saved section count. */
  sectionEndpointRemoval: TrackedSectionEndpointRemoval | null;
};

export type InternalParagraphChangeTrackerState =
  ParagraphChangeTrackerState & {
    /** Edited paragraph positions retained across allocator-only transactions. */
    affectedParagraphPositions: Set<number>;
    /** Editing an already unidentified source paragraph cannot become selective. */
    hasUntrackedSourceChanges: boolean;
    blockStructureFingerprint: string;
    /** The section records themselves, by identity. */
    sectionRecords: ReadonlySet<SectionProperties>;
  };

export const referencesOf = (
  records: readonly SectionProperties[],
): RemovedSectionReference[] => {
  const references: RemovedSectionReference[] = [];
  for (const record of records) {
    for (const { type, rId } of record.headerReferences ?? []) {
      references.push({ part: "header", type, relationshipId: rId });
    }
    for (const { type, rId } of record.footerReferences ?? []) {
      references.push({ part: "footer", type, relationshipId: rId });
    }
  }
  return references;
};

type MarkTrackedSectionEndpointRemovalOptions = {
  sourceDoc: PMNode;
  removedEndpointCount: number;
  removedReferences: readonly RemovedSectionReference[];
};

/** Record only a complete, internally consistent endpoint-removal transition. */
export function markTrackedSectionEndpointRemoval(
  tr: Transaction,
  {
    sourceDoc,
    removedEndpointCount,
    removedReferences,
  }: MarkTrackedSectionEndpointRemovalOptions,
): Transaction {
  const sourceStructure = countDocumentStructure(sourceDoc);
  const sourceParagraphEndpointCount = sourceStructure.sectionEndpoints;
  const expectedStructure = countDocumentStructure(tr.doc);
  const expectedParagraphEndpointCount = expectedStructure.sectionEndpoints;
  if (
    sourceParagraphEndpointCount - expectedParagraphEndpointCount !==
    removedEndpointCount
  ) {
    return tr;
  }
  return tr.setMeta(SECTION_ENDPOINT_REMOVAL_META, {
    type: "tracked-section-endpoint-removal-meta",
    authorization: {
      type: "tracked-section-endpoint-removal",
      sourceParagraphEndpointCount,
      expectedParagraphEndpointCount,
      sourceEndpointFingerprint: sourceStructure.sectionEndpointFingerprint,
      expectedEndpointFingerprint: expectedStructure.sectionEndpointFingerprint,
      removedReferences,
    },
  } satisfies TrackedSectionEndpointRemovalMeta);
}

/**
 * Record that `tr` removed section breaks on purpose: a command whose job is
 * to remove the break (not to delete its paragraph mark, which the tracker
 * reads for itself). Only a plain removal is recorded; a transaction that also
 * adds or replaces a record is left as it is.
 */
export function markSectionBreakRemoval(
  tr: Transaction,
  sourceDoc: PMNode,
): Transaction {
  const before = countDocumentStructure(sourceDoc).sectionRecords;
  const after = countDocumentStructure(tr.doc).sectionRecords;
  const removed = [...before].filter((record) => !after.has(record));
  if (
    removed.length === 0 ||
    [...after].some((record) => !before.has(record))
  ) {
    return tr;
  }
  return markTrackedSectionEndpointRemoval(tr, {
    sourceDoc,
    removedEndpointCount: removed.length,
    removedReferences: referencesOf(removed),
  });
}

/**
 * Create a transaction that clears the change tracker
 */
export function clearTrackedChanges(state: EditorState): Transaction {
  return state.tr.setMeta(paragraphChangeTrackerKey, CLEAR_META);
}

export function ignoreTrackedChanges(tr: Transaction): Transaction {
  return tr.setMeta(paragraphChangeTrackerKey, IGNORE_META);
}

export function markStructuralChange(tr: Transaction): Transaction {
  return tr.setMeta(paragraphChangeTrackerKey, STRUCTURAL_META);
}

type MarkChangedParagraphRangesOptions = {
  ranges: readonly { from: number; to: number }[];
  mappingFrom: number;
  replacesStepMapAt?: number;
};

/** Record mark-only paragraph changes whose replacement step has a granular map. */
export function markChangedParagraphRanges(
  tr: Transaction,
  batch: MarkChangedParagraphRangesOptions,
): Transaction {
  const previous = tr.getMeta(CHANGED_PARAGRAPH_RANGES_META);
  const batches = isChangedParagraphRangesMeta(previous)
    ? [...previous.batches, batch]
    : [batch];
  return tr.setMeta(CHANGED_PARAGRAPH_RANGES_META, {
    type: "changed-paragraph-ranges",
    batches,
  } satisfies ChangedParagraphRangesMeta);
}
