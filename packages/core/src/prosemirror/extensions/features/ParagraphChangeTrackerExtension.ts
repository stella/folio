/**
 * Paragraph Change Tracker Extension
 *
 * Watches ProseMirror transactions and records which paragraph IDs (paraId)
 * were modified. Also detects structural changes (paragraphs added/deleted).
 * Used by the selective save system to patch only changed paragraphs in document.xml.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorState, Transaction } from "prosemirror-state";
import {
  AddMarkStep,
  AddNodeMarkStep,
  AttrStep,
  RemoveMarkStep,
  RemoveNodeMarkStep,
} from "prosemirror-transform";
import type { Mapping } from "prosemirror-transform";

import type {
  RemovedSectionReference,
  TrackedSectionEndpointRemoval,
} from "../../../internal/sectionEndpointResolution";
import { canonicalJson } from "../../../utils/canonicalJson";
import { createExtension } from "../create";
import type { ExtensionRuntime } from "../types";

export const paragraphChangeTrackerKey = new PluginKey<ParagraphChangeTrackerState>(
  "paragraphChangeTracker",
);

const CLEAR_META = "clear";
const IGNORE_META = "ignore";
const STRUCTURAL_META = "structural";
const CHANGED_PARAGRAPH_RANGES_META = "folioChangedParagraphRanges";
const SECTION_ENDPOINT_REMOVAL_META = "folioSectionEndpointRemoval";

type ChangedParagraphRangeBatch = {
  ranges: readonly { from: number; to: number }[];
  mappingFrom: number;
};

type ChangedParagraphRangesMeta = {
  type: "changed-paragraph-ranges";
  batches: readonly ChangedParagraphRangeBatch[];
};

const isChangedParagraphRangesMeta = (value: unknown): value is ChangedParagraphRangesMeta =>
  typeof value === "object" &&
  value !== null &&
  "type" in value &&
  value.type === "changed-paragraph-ranges";

const isRemovedSectionReference = (value: unknown): value is RemovedSectionReference =>
  typeof value === "object" &&
  value !== null &&
  "part" in value &&
  (value.part === "header" || value.part === "footer") &&
  "type" in value &&
  (value.type === "default" || value.type === "first" || value.type === "even") &&
  "relationshipId" in value &&
  typeof value.relationshipId === "string";

const isTrackedSectionEndpointRemoval = (value: unknown): value is TrackedSectionEndpointRemoval =>
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

const isTrackedSectionEndpointRemovalMeta = (
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
  /** Whether paragraphs were added or deleted (structural change) */
  structuralChange: boolean;
  /** Whether any edited paragraph lacked a paraId */
  hasUntrackedChanges: boolean;
  /** Cached paragraph count to avoid full doc traversal on every transaction */
  paragraphCount: number;
  /** Cached section-endpoint count used to invalidate stale save authorization. */
  sectionEndpointCount: number;
  /** Exact endpoint owners and properties used to invalidate count-neutral changes. */
  sectionEndpointFingerprint: string;
  /** Exact tracked-resolution transition that may reduce the saved section count. */
  sectionEndpointRemoval: TrackedSectionEndpointRemoval | null;
};

type DocumentStructureCounts = {
  paragraphs: number;
  sectionEndpoints: number;
  sectionEndpointFingerprint: string;
};

/** Count the structural values the tracker compares for every transaction. */
function countDocumentStructure(doc: PMNode): DocumentStructureCounts {
  let paragraphs = 0;
  let sectionEndpoints = 0;
  const endpointRecords: {
    path: string;
    paraId: unknown;
    sectionBreakType: unknown;
    sectionProperties: unknown;
  }[] = [];
  const visit = (parent: PMNode, parentPath: string): void => {
    parent.forEach((node, _offset, index) => {
      if (node.type.name === "paragraph") {
        paragraphs++;
        if (node.attrs["_sectionProperties"] != null || node.attrs["sectionBreakType"] != null) {
          sectionEndpoints++;
          endpointRecords.push({
            path: parentPath.length === 0 ? `${index}` : `${parentPath}.${index}`,
            paraId: node.attrs["paraId"] ?? null,
            sectionBreakType: node.attrs["sectionBreakType"] ?? null,
            sectionProperties: node.attrs["_sectionProperties"] ?? null,
          });
        }
        return;
      }
      if (node.childCount > 0) {
        const path = parentPath.length === 0 ? `${index}` : `${parentPath}.${index}`;
        visit(node, path);
      }
    });
  };
  visit(doc, "");
  return {
    paragraphs,
    sectionEndpoints,
    sectionEndpointFingerprint: canonicalJson(endpointRecords),
  };
}

/**
 * Collect paraIds of all paragraphs that overlap with the given range
 */
function collectAffectedParaIds(
  doc: EditorState["doc"],
  from: number,
  to: number,
): { ids: Set<string>; hasUntracked: boolean } {
  const ids = new Set<string>();
  let hasUntracked = false;

  doc.nodesBetween(from, to, (node) => {
    if (node.type.name === "paragraph") {
      // SAFETY: paraId is always a string or null per paragraph schema
      const paraId = node.attrs["paraId"] as unknown as string | null;
      if (paraId) {
        ids.add(paraId);
      } else {
        hasUntracked = true;
      }
    }
  });

  return { ids, hasUntracked };
}

/**
 * AddMarkStep / RemoveMarkStep inherit Step.getMap() → StepMap.empty, so we use
 * their from/to to find affected paragraphs.
 * Node mark steps use a single position before the target node.
 */
function collectAffectedParaIdsFromMarkLikeStep(
  doc: EditorState["doc"],
  from: number,
  to: number,
): { ids: Set<string>; hasUntracked: boolean } {
  const lo = Math.min(from, to);
  const hi = Math.max(from, to);
  const end = hi > lo ? hi : lo + 1;
  const primary = collectAffectedParaIds(doc, lo, end);
  if (primary.ids.size > 0 || primary.hasUntracked) {
    return primary;
  }
  // Collapsed range (e.g. empty paragraph): walk up to enclosing paragraph
  try {
    const $p = doc.resolve(lo);
    for (let d = $p.depth; d >= 0; d--) {
      const n = $p.node(d);
      if (n.type.name === "paragraph") {
        // SAFETY: paraId is always a string or null per paragraph schema
        const paraId = n.attrs["paraId"] as unknown as string | null;
        if (paraId) {
          return { ids: new Set([paraId]), hasUntracked: false };
        }
        return { ids: new Set(), hasUntracked: true };
      }
    }
  } catch {
    // ignore
  }
  return { ids: new Set(), hasUntracked: false };
}

function mapStepPosition(remap: Mapping, pos: number, assoc: 1 | -1): number {
  // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map uses assoc as its second parameter.
  return remap.map(pos, assoc);
}

function createParagraphChangeTrackerPlugin(): Plugin<ParagraphChangeTrackerState> {
  return new Plugin<ParagraphChangeTrackerState>({
    key: paragraphChangeTrackerKey,
    state: {
      init(_config, state): ParagraphChangeTrackerState {
        const counts = countDocumentStructure(state.doc);
        return {
          changedParaIds: new Set(),
          structuralChange: false,
          hasUntrackedChanges: false,
          paragraphCount: counts.paragraphs,
          sectionEndpointCount: counts.sectionEndpoints,
          sectionEndpointFingerprint: counts.sectionEndpointFingerprint,
          sectionEndpointRemoval: null,
        };
      },
      apply(tr: Transaction, prevState: ParagraphChangeTrackerState): ParagraphChangeTrackerState {
        const meta = tr.getMeta(paragraphChangeTrackerKey);
        const changedParagraphRangesMeta = tr.getMeta(CHANGED_PARAGRAPH_RANGES_META);
        const sectionEndpointRemovalMeta = tr.getMeta(SECTION_ENDPOINT_REMOVAL_META);
        // Check for explicit clear meta
        if (meta === CLEAR_META) {
          const counts = countDocumentStructure(tr.doc);
          return {
            changedParaIds: new Set(),
            structuralChange: false,
            hasUntrackedChanges: false,
            paragraphCount: counts.paragraphs,
            sectionEndpointCount: counts.sectionEndpoints,
            sectionEndpointFingerprint: counts.sectionEndpointFingerprint,
            sectionEndpointRemoval: null,
          };
        }

        if (meta === IGNORE_META) {
          const counts = tr.docChanged
            ? countDocumentStructure(tr.doc)
            : {
                paragraphs: prevState.paragraphCount,
                sectionEndpoints: prevState.sectionEndpointCount,
                sectionEndpointFingerprint: prevState.sectionEndpointFingerprint,
              };
          return {
            ...prevState,
            paragraphCount: counts.paragraphs,
            sectionEndpointCount: counts.sectionEndpoints,
            sectionEndpointRemoval:
              counts.sectionEndpointFingerprint === prevState.sectionEndpointFingerprint
                ? prevState.sectionEndpointRemoval
                : null,
            sectionEndpointFingerprint: counts.sectionEndpointFingerprint,
          };
        }

        // If no doc changes, keep previous state
        if (!tr.docChanged) {
          return prevState;
        }

        // Count paragraphs in new doc only (use cached count for old doc)
        const counts = countDocumentStructure(tr.doc);
        const newCount = counts.paragraphs;
        let sectionEndpointRemoval = prevState.sectionEndpointRemoval;
        if (counts.sectionEndpointFingerprint !== prevState.sectionEndpointFingerprint) {
          if (
            isTrackedSectionEndpointRemovalMeta(sectionEndpointRemovalMeta) &&
            sectionEndpointRemovalMeta.authorization.sourceEndpointFingerprint ===
              prevState.sectionEndpointFingerprint &&
            sectionEndpointRemovalMeta.authorization.expectedEndpointFingerprint ===
              counts.sectionEndpointFingerprint &&
            sectionEndpointRemovalMeta.authorization.sourceParagraphEndpointCount ===
              prevState.sectionEndpointCount &&
            sectionEndpointRemovalMeta.authorization.expectedParagraphEndpointCount ===
              counts.sectionEndpoints &&
            counts.sectionEndpoints < prevState.sectionEndpointCount
          ) {
            const authorization = sectionEndpointRemovalMeta.authorization;
            sectionEndpointRemoval = {
              type: "tracked-section-endpoint-removal",
              sourceParagraphEndpointCount:
                prevState.sectionEndpointRemoval?.expectedParagraphEndpointCount ===
                authorization.sourceParagraphEndpointCount
                  ? prevState.sectionEndpointRemoval.sourceParagraphEndpointCount
                  : authorization.sourceParagraphEndpointCount,
              expectedParagraphEndpointCount: authorization.expectedParagraphEndpointCount,
              sourceEndpointFingerprint:
                prevState.sectionEndpointRemoval?.expectedParagraphEndpointCount ===
                authorization.sourceParagraphEndpointCount
                  ? prevState.sectionEndpointRemoval.sourceEndpointFingerprint
                  : authorization.sourceEndpointFingerprint,
              expectedEndpointFingerprint: authorization.expectedEndpointFingerprint,
              removedReferences: [
                ...(prevState.sectionEndpointRemoval?.expectedParagraphEndpointCount ===
                authorization.sourceParagraphEndpointCount
                  ? prevState.sectionEndpointRemoval.removedReferences
                  : []),
                ...authorization.removedReferences,
              ],
            };
          } else {
            sectionEndpointRemoval = null;
          }
        }

        // Clone previous state
        const newState: ParagraphChangeTrackerState = {
          changedParaIds: new Set(prevState.changedParaIds),
          structuralChange: prevState.structuralChange || meta === STRUCTURAL_META,
          hasUntrackedChanges: prevState.hasUntrackedChanges,
          paragraphCount: newCount,
          sectionEndpointCount: counts.sectionEndpoints,
          sectionEndpointFingerprint: counts.sectionEndpointFingerprint,
          sectionEndpointRemoval,
        };

        if (isChangedParagraphRangesMeta(changedParagraphRangesMeta)) {
          for (const batch of changedParagraphRangesMeta.batches) {
            const remap = tr.mapping.slice(batch.mappingFrom);
            for (const range of batch.ranges) {
              const from = mapStepPosition(remap, range.from, 1);
              const to = mapStepPosition(remap, range.to, -1);
              if (to <= from) {
                continue;
              }
              const { ids, hasUntracked } = collectAffectedParaIdsFromMarkLikeStep(
                tr.doc,
                from,
                to,
              );
              for (const id of ids) {
                newState.changedParaIds.add(id);
              }
              if (hasUntracked) {
                newState.hasUntrackedChanges = true;
              }
            }
          }
        }

        // Check for structural changes (paragraph count changed)
        if (prevState.paragraphCount !== newCount) {
          newState.structuralChange = true;
        }

        // Track which paragraphs were affected by each step. Step coordinates are
        // valid in the document state where that step ran, so remap them through
        // subsequent steps before reading from the final transaction document.
        for (let stepIndex = 0; stepIndex < tr.steps.length; stepIndex++) {
          // SAFETY: loop condition keeps stepIndex within tr.steps bounds.
          const step = tr.steps[stepIndex]!;
          const remap = tr.mapping.slice(stepIndex + 1);

          if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
            const from = mapStepPosition(remap, step.from, 1);
            const to = mapStepPosition(remap, step.to, -1);
            if (to <= from) {
              continue;
            }
            const { ids, hasUntracked } = collectAffectedParaIdsFromMarkLikeStep(tr.doc, from, to);
            for (const id of ids) {
              newState.changedParaIds.add(id);
            }
            if (hasUntracked) {
              newState.hasUntrackedChanges = true;
            }
            continue;
          }

          // `AttrStep` and the node-mark steps carry a position and an EMPTY
          // step map, so the generic `stepMap.forEach` below never visits
          // them. Reading their position directly is what keeps an
          // attribute-only edit — a paragraph mark on a blank paragraph, a
          // list level, a style id — from saving as the original XML.
          if (
            step instanceof AddNodeMarkStep ||
            step instanceof RemoveNodeMarkStep ||
            step instanceof AttrStep
          ) {
            const pos = mapStepPosition(remap, step.pos, 1);
            const node = tr.doc.nodeAt(pos);
            if (!node) {
              continue;
            }
            const end = pos + node.nodeSize;
            const { ids, hasUntracked } = collectAffectedParaIds(tr.doc, pos, end);
            for (const id of ids) {
              newState.changedParaIds.add(id);
            }
            if (hasUntracked) {
              newState.hasUntrackedChanges = true;
            }
            continue;
          }

          const stepMap = step.getMap();
          // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror StepMap.forEach
          stepMap.forEach((_oldStart, _oldEnd, newStart, newEnd) => {
            const from = mapStepPosition(remap, newStart, 1);
            const to = mapStepPosition(remap, newEnd, -1);
            if (to < from) {
              return;
            }
            const { ids, hasUntracked } = collectAffectedParaIds(tr.doc, from, to);
            for (const id of ids) {
              newState.changedParaIds.add(id);
            }
            if (hasUntracked) {
              newState.hasUntrackedChanges = true;
            }
          });
        }

        return newState;
      },
    },
  });
}

/**
 * Get the change tracker state from an EditorState
 */
export function getChangeTrackerState(state: EditorState): ParagraphChangeTrackerState | undefined {
  return paragraphChangeTrackerKey.getState(state);
}

/**
 * Get the set of changed paragraph IDs from an EditorState
 */
export function getChangedParagraphIds(state: EditorState): Set<string> {
  return getChangeTrackerState(state)?.changedParaIds ?? new Set();
}

/**
 * Check if structural changes (paragraph add/delete) occurred
 */
export function hasStructuralChanges(state: EditorState): boolean {
  const trackerState = getChangeTrackerState(state);
  return trackerState?.structuralChange ?? false;
}

/**
 * Check if any changes affected paragraphs without paraId
 */
export function hasUntrackedChanges(state: EditorState): boolean {
  const trackerState = getChangeTrackerState(state);
  return trackerState?.hasUntrackedChanges ?? false;
}

/** Exact section-count transition authorized by tracked paragraph-mark resolution. */
export function getTrackedSectionEndpointRemoval(
  state: EditorState,
): TrackedSectionEndpointRemoval | null {
  return getChangeTrackerState(state)?.sectionEndpointRemoval ?? null;
}

type MarkTrackedSectionEndpointRemovalOptions = {
  sourceDoc: PMNode;
  removedEndpointCount: number;
  removedReferences: readonly RemovedSectionReference[];
};

/** Record only a complete, internally consistent endpoint-removal transition. */
export function markTrackedSectionEndpointRemoval(
  tr: Transaction,
  { sourceDoc, removedEndpointCount, removedReferences }: MarkTrackedSectionEndpointRemovalOptions,
): Transaction {
  const sourceStructure = countDocumentStructure(sourceDoc);
  const sourceParagraphEndpointCount = sourceStructure.sectionEndpoints;
  const expectedStructure = countDocumentStructure(tr.doc);
  const expectedParagraphEndpointCount = expectedStructure.sectionEndpoints;
  if (sourceParagraphEndpointCount - expectedParagraphEndpointCount !== removedEndpointCount) {
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
};

/** Record mark-only paragraph changes whose replacement step has a granular map. */
export function markChangedParagraphRanges(
  tr: Transaction,
  batch: MarkChangedParagraphRangesOptions,
): Transaction {
  const previous = tr.getMeta(CHANGED_PARAGRAPH_RANGES_META);
  const batches = isChangedParagraphRangesMeta(previous) ? [...previous.batches, batch] : [batch];
  return tr.setMeta(CHANGED_PARAGRAPH_RANGES_META, {
    type: "changed-paragraph-ranges",
    batches,
  } satisfies ChangedParagraphRangesMeta);
}

export const ParagraphChangeTrackerExtension = createExtension({
  name: "paragraphChangeTracker",
  defaultOptions: {},
  onSchemaReady(): ExtensionRuntime {
    return {
      plugins: [createParagraphChangeTrackerPlugin()],
    };
  },
});
