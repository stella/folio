import {
  parseFolioDocumentOperationBatch,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
} from "../document-operations";
import { getFolioParaIdFromBlockId } from "../types/block-id";
import { canonicalJson } from "../utils/canonicalJson";
import { getDocumentStyleResolver } from "../prosemirror/plugins/documentStyles";
import { rejectAllSuggestions } from "../prosemirror/commands/comments";
import type { EditorState } from "prosemirror-state";
import {
  createFolioAIEditSnapshotWithStyleResolver,
  hashFolioAIBlockText,
  sourceDocumentOf,
} from "./snapshot";
import type { FolioDocumentStoryHandle } from "./headless";
import type {
  FolioAIEditAppliedOperation,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
} from "./types";

export const FOLIO_PENDING_SUGGESTION_VERSION = 1 as const;

/** A host-owned, JSON-serializable proposal. It is never written into the DOCX. */
export type FolioPendingSuggestionRecord = {
  version: typeof FOLIO_PENDING_SUGGESTION_VERSION;
  suggestionId: string;
  operation: FolioAIEditOperation;
  story: FolioDocumentStoryHandle;
  anchor: {
    blockId: string;
    paraId: string | null;
    originalTextHash: string;
    selectedTextHash?: string;
    startOffset?: number;
    endOffset?: number;
  };
  author: string;
  provenance: "suggested";
  sourceDocumentFingerprint: string;
  commentId?: number;
};

export type FolioPendingSuggestionStaleReason =
  | "missingAnchor"
  | "ambiguousAnchor"
  | "textChanged"
  | "unsupportedVersion"
  | "documentChanged"
  | "invalidRecord"
  | "applyFailed";

export type FolioPendingSuggestionLoadResult =
  | { status: "restaged"; suggestionId: string }
  | {
      status: "stale";
      suggestionId: string | null;
      reason: FolioPendingSuggestionStaleReason;
    };

type RecordAppliedOptions = {
  snapshot: FolioAIEditSnapshot;
  story: FolioDocumentStoryHandle;
  operations: readonly FolioAIEditOperation[];
  applied: readonly FolioAIEditAppliedOperation[];
  author: string;
  activeSuggestionIds: ReadonlySet<string>;
};

type ReplaceRestagedStoryOptions = {
  story: FolioDocumentStoryHandle;
  restaged: readonly {
    snapshot: FolioAIEditSnapshot;
    operation: FolioAIEditOperation;
    applied: FolioAIEditAppliedOperation;
    author: string;
  }[];
  activeSuggestionIds: ReadonlySet<string>;
};

type LoadPendingSuggestionsOptions = {
  records: readonly unknown[];
  snapshotForStory: (story: FolioDocumentStoryHandle) => FolioAIEditSnapshot | null;
  sourceSnapshotForStory: (
    story: FolioDocumentStoryHandle,
    commentIds: ReadonlySet<number>,
  ) => FolioAIEditSnapshot | null;
  apply: (
    record: FolioPendingSuggestionRecord,
    snapshot: FolioAIEditSnapshot,
  ) => FolioAIEditApplyResult;
  activeSuggestionIds: () => ReadonlySet<string>;
};

type ExportPendingSuggestionsOptions = {
  activeSuggestionIds: ReadonlySet<string>;
  snapshotForStory: (
    story: FolioDocumentStoryHandle,
    commentIds: ReadonlySet<number>,
  ) => FolioAIEditSnapshot | null;
};

/** Drop only comments attached to proposals that have not been accepted. */
export const withoutPendingSuggestionCommentMarks = (
  state: EditorState,
  commentIds: ReadonlySet<number>,
): EditorState => {
  if (commentIds.size === 0) return state;
  let transaction = state.tr;
  state.doc.descendants((node, position) => {
    for (const mark of node.marks) {
      if (mark.type.name === "comment" && commentIds.has(mark.attrs["commentId"])) {
        transaction = transaction.removeMark(position, position + node.nodeSize, mark);
      }
    }
  });
  return transaction.docChanged ? state.apply(transaction) : state;
};

/** The saved package's story, with all pending proposals removed. */
export const createPendingSuggestionSourceSnapshot = (
  state: EditorState,
  commentIds: ReadonlySet<number>,
): FolioAIEditSnapshot => {
  let source = state;
  rejectAllSuggestions()(state, (transaction) => {
    source = state.apply(transaction);
  });
  source = withoutPendingSuggestionCommentMarks(source, commentIds);
  return createFolioAIEditSnapshotWithStyleResolver(source.doc, getDocumentStyleResolver(source));
};

const isObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const operationBlockId = (operation: FolioAIEditOperation): string => {
  switch (operation.type) {
    case "replaceRange":
    case "commentOnRange":
    case "formatRange":
      return operation.range.blockId;
    default:
      return operation.blockId;
  }
};

const storyFromUnknown = (value: unknown): FolioDocumentStoryHandle | null => {
  if (!isObject(value)) return null;
  switch (value["type"]) {
    case "main":
      return { type: "main" };
    case "header":
    case "footer":
      return typeof value["relationshipId"] === "string"
        ? { type: value["type"], relationshipId: value["relationshipId"] }
        : null;
    case "footnote":
    case "endnote":
      return typeof value["noteId"] === "number" && Number.isSafeInteger(value["noteId"])
        ? { type: value["type"], noteId: value["noteId"] }
        : null;
    default:
      return null;
  }
};

const fingerprintOf = (snapshot: FolioAIEditSnapshot, story: FolioDocumentStoryHandle): string =>
  hashFolioAIBlockText(
    canonicalJson([
      story,
      snapshot.blocks.map((block) => [block, snapshot.anchors[block.id]?.structuralBoundaryHash]),
    ]),
  );

const paraIdCountsOf = (snapshot: FolioAIEditSnapshot): Map<string, number> => {
  const counts = new Map<string, number>();
  sourceDocumentOf(snapshot).descendants((node) => {
    const paraId = node.attrs["paraId"];
    if (node.type.name === "paragraph" && typeof paraId === "string") {
      counts.set(paraId, (counts.get(paraId) ?? 0) + 1);
    }
  });
  return counts;
};

const recordKey = (record: FolioPendingSuggestionRecord): string =>
  canonicalJson([record.story, record.suggestionId, record.operation.id]);

const parseRecord = (value: unknown): FolioPendingSuggestionRecord | null => {
  if (!isObject(value) || !isObject(value["anchor"])) return null;
  const story = storyFromUnknown(value["story"]);
  const anchor = value["anchor"];
  if (
    story === null ||
    typeof value["suggestionId"] !== "string" ||
    typeof value["author"] !== "string" ||
    value["provenance"] !== "suggested" ||
    typeof value["sourceDocumentFingerprint"] !== "string" ||
    (value["commentId"] !== undefined &&
      (typeof value["commentId"] !== "number" || !Number.isSafeInteger(value["commentId"]))) ||
    typeof anchor["blockId"] !== "string" ||
    (anchor["paraId"] !== null && typeof anchor["paraId"] !== "string") ||
    typeof anchor["originalTextHash"] !== "string"
  ) {
    return null;
  }
  let operation: FolioAIEditOperation;
  try {
    const batch = parseFolioDocumentOperationBatch({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: [value["operation"]],
      mode: "suggested",
    });
    const parsed = batch.operations.at(0);
    if (!parsed) return null;
    operation = parsed;
  } catch {
    // Persisted JSON is an external input boundary; a malformed operation is inert.
    return null;
  }
  if (
    operationBlockId(operation) !== anchor["blockId"] ||
    (operation.suggestionId ?? operation.id) !== value["suggestionId"] ||
    getFolioParaIdFromBlockId(anchor["blockId"]) !== anchor["paraId"]
  ) {
    return null;
  }
  if (operation.type === "replaceRange" || operation.type === "formatRange") {
    if (
      anchor["selectedTextHash"] !== operation.range.selectedTextHash ||
      anchor["startOffset"] !== operation.range.startOffset ||
      anchor["endOffset"] !== operation.range.endOffset
    )
      return null;
  }
  return {
    version: FOLIO_PENDING_SUGGESTION_VERSION,
    suggestionId: value["suggestionId"],
    operation,
    story,
    anchor: {
      blockId: anchor["blockId"],
      paraId: anchor["paraId"],
      originalTextHash: anchor["originalTextHash"],
      ...(typeof anchor["selectedTextHash"] === "string" && {
        selectedTextHash: anchor["selectedTextHash"],
      }),
      ...(typeof anchor["startOffset"] === "number" && {
        startOffset: anchor["startOffset"],
      }),
      ...(typeof anchor["endOffset"] === "number" && { endOffset: anchor["endOffset"] }),
    },
    author: value["author"],
    provenance: "suggested",
    sourceDocumentFingerprint: value["sourceDocumentFingerprint"],
    ...(typeof value["commentId"] === "number" &&
      Number.isSafeInteger(value["commentId"]) && { commentId: value["commentId"] }),
  };
};

/** Tracks only operations that successfully staged in `suggested` mode. */
export class FolioPendingSuggestionRegistry {
  private readonly records = new Map<string, FolioPendingSuggestionRecord>();
  private readonly sourceFingerprints = new Map<string, string>();

  clear(): void {
    this.records.clear();
    this.sourceFingerprints.clear();
  }

  recordApplied({
    snapshot,
    story,
    operations,
    applied,
    author,
    activeSuggestionIds,
  }: RecordAppliedOptions): void {
    if (applied.length === 0) return;
    const storyKey = canonicalJson(story);
    for (const [key, record] of this.records) {
      if (!activeSuggestionIds.has(record.suggestionId)) this.records.delete(key);
    }
    if (![...this.records.values()].some((record) => canonicalJson(record.story) === storyKey)) {
      this.sourceFingerprints.delete(storyKey);
    }
    const fingerprint = this.sourceFingerprints.get(storyKey) ?? fingerprintOf(snapshot, story);
    this.sourceFingerprints.set(storyKey, fingerprint);
    const appliedById = new Map(applied.map((entry) => [entry.id, entry]));
    for (const operation of operations) {
      const outcome = appliedById.get(operation.id);
      if (!outcome?.suggestionId) continue;
      const blockId = operationBlockId(operation);
      const anchor = snapshot.anchors[blockId];
      if (!anchor) continue;
      const range =
        operation.type === "replaceRange" || operation.type === "formatRange"
          ? operation.range
          : null;
      const record: FolioPendingSuggestionRecord = {
        version: FOLIO_PENDING_SUGGESTION_VERSION,
        suggestionId: outcome.suggestionId,
        operation: structuredClone(operation),
        story: structuredClone(story),
        anchor: {
          blockId,
          paraId: getFolioParaIdFromBlockId(blockId),
          originalTextHash: anchor.textHash,
          ...(range !== null && {
            selectedTextHash: range.selectedTextHash,
            startOffset: range.startOffset,
            endOffset: range.endOffset,
          }),
        },
        author,
        provenance: "suggested",
        sourceDocumentFingerprint: fingerprint,
        ...(outcome.commentId !== undefined && { commentId: outcome.commentId }),
      };
      this.records.set(recordKey(record), record);
    }
  }

  /** Rebase records after ordinary revisions change their source story. */
  replaceRestagedStory({
    story,
    restaged,
    activeSuggestionIds,
  }: ReplaceRestagedStoryOptions): void {
    if (restaged.length === 0) return;
    const storyKey = canonicalJson(story);
    for (const [key, record] of this.records) {
      if (canonicalJson(record.story) === storyKey) this.records.delete(key);
    }
    const first = restaged.at(0);
    if (!first) return;
    this.sourceFingerprints.set(storyKey, fingerprintOf(first.snapshot, story));
    for (const { snapshot, operation, applied, author } of restaged) {
      this.recordApplied({
        snapshot,
        story,
        operations: [operation],
        applied: [applied],
        author,
        activeSuggestionIds,
      });
    }
  }

  exportPendingSuggestions({
    activeSuggestionIds,
    snapshotForStory,
  }: ExportPendingSuggestionsOptions): FolioPendingSuggestionRecord[] {
    const pending: FolioPendingSuggestionRecord[] = [];
    const fingerprints = new Map<string, string>();
    const commentIdsByStory = new Map<string, Set<number>>();
    for (const record of this.records.values()) {
      if (!activeSuggestionIds.has(record.suggestionId) || record.commentId === undefined) continue;
      const storyKey = canonicalJson(record.story);
      const commentIds = commentIdsByStory.get(storyKey) ?? new Set<number>();
      commentIds.add(record.commentId);
      commentIdsByStory.set(storyKey, commentIds);
    }
    for (const [key, record] of this.records) {
      if (activeSuggestionIds.has(record.suggestionId)) {
        const storyKey = canonicalJson(record.story);
        let fingerprint = fingerprints.get(storyKey);
        if (fingerprint === undefined) {
          const snapshot = snapshotForStory(
            record.story,
            commentIdsByStory.get(storyKey) ?? new Set<number>(),
          );
          if (!snapshot) continue;
          fingerprint = fingerprintOf(snapshot, record.story);
          fingerprints.set(storyKey, fingerprint);
        }
        const current = { ...record, sourceDocumentFingerprint: fingerprint };
        this.records.set(key, current);
        pending.push(structuredClone(current));
      } else {
        this.records.delete(key);
      }
    }
    for (const [storyKey, fingerprint] of fingerprints) {
      this.sourceFingerprints.set(storyKey, fingerprint);
    }
    if (pending.length === 0) this.sourceFingerprints.clear();
    return pending;
  }

  loadPendingSuggestions({
    records,
    snapshotForStory,
    sourceSnapshotForStory,
    apply,
    activeSuggestionIds,
  }: LoadPendingSuggestionsOptions): FolioPendingSuggestionLoadResult[] {
    const baselines = new Map<
      string,
      { snapshot: FolioAIEditSnapshot; fingerprint: string } | null
    >();
    const changedStories = new Set<string>();
    return records.map((value): FolioPendingSuggestionLoadResult => {
      const suggestionId =
        isObject(value) && typeof value["suggestionId"] === "string" ? value["suggestionId"] : null;
      if (!isObject(value) || value["version"] !== FOLIO_PENDING_SUGGESTION_VERSION) {
        return { status: "stale", suggestionId, reason: "unsupportedVersion" };
      }
      const record = parseRecord(value);
      if (!record) return { status: "stale", suggestionId, reason: "invalidRecord" };
      const key = recordKey(record);
      const existing = this.records.get(key);
      if (existing && activeSuggestionIds().has(record.suggestionId)) {
        const { commentId: _existingCommentId, ...existingProposal } = existing;
        const { commentId: _loadedCommentId, ...loadedProposal } = record;
        return canonicalJson(existingProposal) === canonicalJson(loadedProposal)
          ? { status: "restaged", suggestionId: record.suggestionId }
          : { status: "stale", suggestionId, reason: "invalidRecord" };
      }
      const storyKey = canonicalJson(record.story);
      if (!baselines.has(storyKey)) {
        const snapshot = snapshotForStory(record.story);
        const active = activeSuggestionIds();
        const commentIds = new Set<number>();
        for (const existingRecord of this.records.values()) {
          if (
            canonicalJson(existingRecord.story) === storyKey &&
            active.has(existingRecord.suggestionId) &&
            existingRecord.commentId !== undefined
          ) {
            commentIds.add(existingRecord.commentId);
          }
        }
        // Fingerprints describe the saved story; anchors describe the live
        // story, where earlier records may already have been restaged.
        const source = sourceSnapshotForStory(record.story, commentIds);
        baselines.set(
          storyKey,
          snapshot && source
            ? {
                snapshot,
                fingerprint: fingerprintOf(source, record.story),
              }
            : null,
        );
      }
      const baseline = baselines.get(storyKey);
      if (!baseline || record.anchor.paraId === null) {
        return { status: "stale", suggestionId, reason: "missingAnchor" };
      }
      const snapshot = changedStories.has(storyKey)
        ? snapshotForStory(record.story)
        : baseline.snapshot;
      if (!snapshot) return { status: "stale", suggestionId, reason: "missingAnchor" };
      const multiplicity = paraIdCountsOf(snapshot).get(record.anchor.paraId) ?? 0;
      if (multiplicity === 0 || !snapshot.anchors[record.anchor.paraId]) {
        return { status: "stale", suggestionId, reason: "missingAnchor" };
      }
      if (multiplicity > 1) {
        return { status: "stale", suggestionId, reason: "ambiguousAnchor" };
      }
      const anchor = snapshot.anchors[record.anchor.paraId];
      if (!anchor || anchor.textHash !== record.anchor.originalTextHash) {
        return { status: "stale", suggestionId, reason: "textChanged" };
      }
      const { startOffset, endOffset, selectedTextHash } = record.anchor;
      if (
        startOffset !== undefined &&
        endOffset !== undefined &&
        selectedTextHash !== undefined &&
        hashFolioAIBlockText(anchor.text.slice(startOffset, endOffset)) !== selectedTextHash
      ) {
        return { status: "stale", suggestionId, reason: "textChanged" };
      }
      if (baseline.fingerprint !== record.sourceDocumentFingerprint) {
        return { status: "stale", suggestionId, reason: "documentChanged" };
      }
      const outcome = apply(record, snapshot);
      const applied = outcome.applied.find((entry) => entry.id === record.operation.id);
      if (!applied) {
        return { status: "stale", suggestionId, reason: "applyFailed" };
      }
      this.sourceFingerprints.set(storyKey, record.sourceDocumentFingerprint);
      if (!this.records.has(key)) {
        const { commentId: _oldCommentId, ...rest } = record;
        this.records.set(
          key,
          structuredClone({
            ...rest,
            ...(applied.commentId !== undefined && { commentId: applied.commentId }),
          }),
        );
      }
      changedStories.add(storyKey);
      return { status: "restaged", suggestionId: record.suggestionId };
    });
  }
}
