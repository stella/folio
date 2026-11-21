import { panic } from "better-result";
import type { EditorState } from "prosemirror-state";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  parseFolioDocumentOperationBatch,
} from "../document-operations";
import {
  getSuggestions,
  rejectEverySuggestion,
} from "../prosemirror/commands/comments";
import { getDocumentStyleResolver } from "../prosemirror/plugins/documentStyleState";
import { getFolioParaIdFromBlockId } from "../types/block-id";
import { canonicalJson } from "../utils/canonicalJson";
import { persistedBlock } from "./block-identity";
import type { FolioDocumentStoryHandle } from "./headless";
import {
  createFolioAIEditSnapshotWithStyleResolver,
  hashFolioAIBlockText,
  sourceDocumentOf,
} from "./snapshot";
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
  snapshotForStory: (
    story: FolioDocumentStoryHandle,
  ) => FolioAIEditSnapshot | null;
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

type PrepareAbsorptionOptions = {
  /** Absorbed paragraph id -> the paragraph that holds its content now. */
  survivors: ReadonlyMap<string, string>;
  activeSuggestionIds: ReadonlySet<string>;
  sourceSnapshotForStory: (
    story: FolioDocumentStoryHandle,
  ) => FolioAIEditSnapshot | null;
};

type AbsorptionStep = (
  sourceSnapshotForStory: (
    story: FolioDocumentStoryHandle,
  ) => FolioAIEditSnapshot | null,
) => void;

type PlaceRemovedAnchorsOptions = {
  activeSuggestionIds: ReadonlySet<string>;
  stateForStory: (story: FolioDocumentStoryHandle) => EditorState | null;
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
      if (
        mark.type.name === "comment" &&
        commentIds.has(mark.attrs["commentId"])
      ) {
        transaction = transaction.removeMark(
          position,
          position + node.nodeSize,
          mark,
        );
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
  const source = withoutPendingSuggestionCommentMarks(
    rejectEverySuggestion(state),
    commentIds,
  );
  return createFolioAIEditSnapshotWithStyleResolver(
    source.doc,
    getDocumentStyleResolver(source),
  );
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

/** `operation` addressed to `blockId`, its text range moved by `shift`. */
const operationOnBlock = (
  operation: FolioAIEditOperation,
  blockId: string,
  shift: number,
): FolioAIEditOperation => {
  switch (operation.type) {
    case "replaceRange":
    case "commentOnRange":
    case "formatRange":
      return {
        ...operation,
        range: {
          ...operation.range,
          blockId,
          startOffset: operation.range.startOffset + shift,
          endOffset: operation.range.endOffset + shift,
        },
      };
    default:
      return { ...operation, blockId };
  }
};

/** The paragraph `paraId` ends in once every absorption is followed. */
const finalSurvivor = (
  survivors: ReadonlyMap<string, string>,
  paraId: string,
): string => {
  let current = paraId;
  const seen = new Set<string>();
  for (
    let next = survivors.get(current);
    next !== undefined;
    next = survivors.get(current)
  ) {
    if (seen.has(next))
      return panic("Paragraph absorptions form a cycle", { paraId });
    seen.add(current);
    current = next;
  }
  return current;
};

type ReanchoredRecordOptions = {
  record: FolioPendingSuggestionRecord;
  survivors: ReadonlyMap<string, string>;
  /** The record's story before the resolution, every proposal rejected. */
  before: FolioAIEditSnapshot | null;
  /** The same story after it. */
  after: FolioAIEditSnapshot | null;
};

/**
 * `record` anchored to the paragraph that absorbed its own. The survivor
 * shows its members' text in document order (an absorption joins paragraphs
 * and keeps every word they show), so a text range moves by the length of the
 * members before the record's paragraph. A survivor showing other text is a
 * resolution this rule does not describe: that panics.
 *
 * `null` when nothing of the record's paragraph lives in a saved paragraph:
 * it showed no text (a paragraph removed whole), or the survivor is a
 * paragraph only a proposal adds. {@link placedRecord} then anchors it.
 */
const reanchoredRecord = ({
  record,
  survivors,
  before,
  after,
}: ReanchoredRecordOptions): FolioPendingSuggestionRecord | null => {
  const { suggestionId, anchor: recorded } = record;
  const { paraId } = recorded;
  if (paraId === null)
    return panic("Only a record with a paragraph id is absorbed", {
      suggestionId,
    });
  const survivor = finalSurvivor(survivors, paraId);
  let shift: number | null = null;
  let text = "";
  for (const { id } of before?.blocks ?? []) {
    if (
      id !== survivor &&
      !(survivors.has(id) && finalSurvivor(survivors, id) === survivor)
    )
      continue;
    if (id === paraId) shift = text.length;
    text += before?.anchors[id]?.text ?? "";
  }
  const anchor = after?.anchors[survivor];
  if (!anchor) return null;
  // The survivor's own record moves by the text absorbed before it, when the
  // survivor shows its members' text; it stays where it is when what joined
  // it is a paragraph only a proposal adds.
  if (paraId === survivor && (shift === null || anchor.text !== text))
    return null;
  if (paraId !== survivor && !before?.anchors[paraId]?.text) return null;
  if (shift === null || anchor.text !== text)
    return panic(
      "A pending suggestion's paragraph was absorbed, but the survivor does not show its text",
      { suggestionId, paraId, survivor },
    );
  return {
    ...record,
    operation: operationOnBlock(record.operation, survivor, shift),
    anchor: {
      blockId: survivor,
      paraId: survivor,
      originalTextHash: anchor.textHash,
      ...(recorded.selectedTextHash !== undefined && {
        selectedTextHash: recorded.selectedTextHash,
      }),
      ...(recorded.startOffset !== undefined && {
        startOffset: recorded.startOffset + shift,
      }),
      ...(recorded.endOffset !== undefined && {
        endOffset: recorded.endOffset + shift,
      }),
    },
  };
};

/** The block list `block` stands in: its containers, and its table cell. */
const containerKey = ({
  containerPath = [],
  table,
}: FolioAIEditSnapshot["blocks"][number]): string =>
  canonicalJson({
    containerPath,
    cell:
      table === undefined
        ? null
        : [table.tableIndex, table.rowIndex, table.cellIndex],
  });

type PlacedRecordOptions = {
  record: FolioPendingSuggestionRecord;
  /** The record's story as it is, every proposal shown. */
  live: FolioAIEditSnapshot;
  /** The same story, every proposal rejected: the paragraphs a save holds. */
  saved: FolioAIEditSnapshot;
  /** The document ranges the record's suggestion covers in `live`. */
  ranges: readonly { from: number; to: number }[];
};

/**
 * `record`, a block insertion whose anchor paragraph is gone, anchored where
 * its paragraphs stand: before the next saved paragraph beside them, or,
 * when none follows, after the previous one. A record of another kind, or
 * one with no saved paragraph beside its own, is a state the rules that
 * remove paragraphs never leave: that panics.
 */
const placedRecord = ({
  record,
  live,
  saved,
  ranges,
}: PlacedRecordOptions): FolioPendingSuggestionRecord => {
  const { operation, suggestionId } = record;
  if (
    operation.type !== "insertAfterBlock" &&
    operation.type !== "insertBeforeBlock"
  )
    return panic(
      "A pending suggestion's paragraph is gone, and only a block insertion keeps a place without it",
      { suggestionId, type: operation.type },
    );
  const savedIds = new Set(saved.blocks.map(({ id }) => id));
  // The first paragraph the suggestion adds.
  const index = live.blocks.findIndex(({ id }) => {
    const anchor = live.anchors[id];
    return (
      !savedIds.has(id) &&
      anchor !== undefined &&
      ranges.some(({ from, to }) => from <= anchor.to && anchor.from <= to)
    );
  });
  const placed = live.blocks.at(index);
  if (index < 0 || placed === undefined)
    return panic("A pending block insertion shows no paragraph of its own", {
      suggestionId,
    });
  const container = containerKey(placed);
  // The nearest saved paragraph in `step`'s direction, if it stands in the
  // same container: a table or another container between them is no place
  // beside it.
  const besideIn = (step: 1 | -1) => {
    for (
      let at = index + step;
      at >= 0 && at < live.blocks.length;
      at += step
    ) {
      const block = live.blocks.at(at);
      if (block === undefined || !savedIds.has(block.id)) continue;
      return containerKey(block) === container ? block : undefined;
    }
    return undefined;
  };
  const next = besideIn(1);
  const neighbour = next ?? besideIn(-1);
  const anchor =
    neighbour === undefined ? undefined : saved.anchors[neighbour.id];
  if (neighbour === undefined || anchor === undefined)
    return panic(
      "A pending block insertion whose anchor is gone has no saved paragraph beside it",
      { suggestionId },
    );
  return {
    ...record,
    operation: {
      ...operation,
      type: next === undefined ? "insertAfterBlock" : "insertBeforeBlock",
      blockId: neighbour.id,
    },
    anchor: {
      blockId: neighbour.id,
      paraId: neighbour.id,
      originalTextHash: anchor.textHash,
    },
  };
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
      return typeof value["noteId"] === "number" &&
        Number.isSafeInteger(value["noteId"])
        ? { type: value["type"], noteId: value["noteId"] }
        : null;
    default:
      return null;
  }
};

const fingerprintOf = (
  snapshot: FolioAIEditSnapshot,
  story: FolioDocumentStoryHandle,
): string =>
  hashFolioAIBlockText(
    canonicalJson([
      story,
      snapshot.blocks.map((block) => [
        persistedBlock(block),
        snapshot.anchors[block.id]?.structuralBoundaryHash,
      ]),
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
      (typeof value["commentId"] !== "number" ||
        !Number.isSafeInteger(value["commentId"]))) ||
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
  if (
    operation.type === "replaceRange" ||
    operation.type === "formatRange" ||
    operation.type === "commentOnRange"
  ) {
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
      ...(typeof anchor["endOffset"] === "number" && {
        endOffset: anchor["endOffset"],
      }),
    },
    author: value["author"],
    provenance: "suggested",
    sourceDocumentFingerprint: value["sourceDocumentFingerprint"],
    ...(typeof value["commentId"] === "number" &&
      Number.isSafeInteger(value["commentId"]) && {
        commentId: value["commentId"],
      }),
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

  /** An independent registry holding the same records. */
  clone(): FolioPendingSuggestionRegistry {
    const copy = new FolioPendingSuggestionRegistry();
    for (const [key, record] of this.records)
      copy.records.set(key, structuredClone(record));
    for (const [key, fingerprint] of this.sourceFingerprints)
      copy.sourceFingerprints.set(key, fingerprint);
    return copy;
  }

  /** Every suggestion id a record names. */
  suggestionIds(): Set<string> {
    return new Set(
      [...this.records.values()].map(({ suggestionId }) => suggestionId),
    );
  }

  /**
   * Anchor every live record whose paragraph is gone from its story, a
   * resolution or an edit having removed it, beside the paragraphs its
   * suggestion added ({@link placedRecord}): a record never names a removed
   * or absorbed paragraph. Run after the change publishes.
   */
  placeRemovedAnchors({
    activeSuggestionIds,
    stateForStory,
  }: PlaceRemovedAnchorsOptions): void {
    const views = new Map<
      string,
      { state: EditorState; live: FolioAIEditSnapshot } | null
    >();
    for (const [key, record] of this.records) {
      const { suggestionId, story, anchor } = record;
      if (!activeSuggestionIds.has(suggestionId) || anchor.paraId === null)
        continue;
      const storyKey = canonicalJson(story);
      let view = views.get(storyKey);
      if (view === undefined) {
        const state = stateForStory(story);
        view =
          state === null
            ? null
            : {
                state,
                live: createFolioAIEditSnapshotWithStyleResolver(
                  state.doc,
                  getDocumentStyleResolver(state),
                ),
              };
        views.set(storyKey, view);
      }
      if (view === null || view.live.anchors[anchor.paraId] !== undefined)
        continue;
      const ranges =
        getSuggestions(view.state).find(
          (entry) => entry.suggestionId === suggestionId,
        )?.ranges ?? [];
      this.records.set(
        key,
        placedRecord({
          record,
          live: view.live,
          saved: createPendingSuggestionSourceSnapshot(view.state, new Set()),
          ranges,
        }),
      );
    }
  }

  /**
   * Prepare re-anchoring for a resolution that absorbs paragraphs: a live
   * proposal anchored to an absorbed paragraph now lives in the survivor, so
   * its record must name the survivor, never the absorbed id; one anchored to
   * the survivor itself moves by the text absorbed before it. Call before the
   * resolution publishes, with snapshots of the story as it is; run the
   * returned step after, with snapshots as the resolution leaves it. `null`
   * when no live record is anchored to an absorbed paragraph.
   */
  prepareAbsorption({
    survivors,
    activeSuggestionIds,
    sourceSnapshotForStory,
  }: PrepareAbsorptionOptions): AbsorptionStep | null {
    // A survivor's own records move too: the text absorbed before it shifts
    // their ranges, and their paragraph's text changes.
    const absorbing = new Set(
      [...survivors.keys()].map((paraId) => finalSurvivor(survivors, paraId)),
    );
    const affected = [...this.records].filter(
      ([, { suggestionId, anchor }]) =>
        activeSuggestionIds.has(suggestionId) &&
        anchor.paraId !== null &&
        (survivors.has(anchor.paraId) || absorbing.has(anchor.paraId)),
    );
    if (affected.length === 0) return null;
    const before = new Map<string, FolioAIEditSnapshot | null>();
    for (const [, { story }] of affected) {
      const storyKey = canonicalJson(story);
      if (!before.has(storyKey))
        before.set(storyKey, sourceSnapshotForStory(story));
    }
    return (afterSnapshotForStory) => {
      const after = new Map<string, FolioAIEditSnapshot | null>();
      for (const [key, record] of affected) {
        const storyKey = canonicalJson(record.story);
        if (!after.has(storyKey))
          after.set(storyKey, afterSnapshotForStory(record.story));
        const reanchored = reanchoredRecord({
          record,
          survivors,
          before: before.get(storyKey) ?? null,
          after: after.get(storyKey) ?? null,
        });
        if (reanchored) this.records.set(key, reanchored);
      }
    };
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
      if (!activeSuggestionIds.has(record.suggestionId))
        this.records.delete(key);
    }
    if (
      ![...this.records.values()].some(
        (record) => canonicalJson(record.story) === storyKey,
      )
    ) {
      this.sourceFingerprints.delete(storyKey);
    }
    const fingerprint =
      this.sourceFingerprints.get(storyKey) ?? fingerprintOf(snapshot, story);
    this.sourceFingerprints.set(storyKey, fingerprint);
    const appliedById = new Map(applied.map((entry) => [entry.id, entry]));
    for (const operation of operations) {
      const outcome = appliedById.get(operation.id);
      if (!outcome?.suggestionId) continue;
      const blockId = operationBlockId(operation);
      const anchor = snapshot.anchors[blockId];
      if (!anchor) continue;
      const range =
        operation.type === "replaceRange" ||
        operation.type === "formatRange" ||
        operation.type === "commentOnRange"
          ? operation.range
          : null;
      // An operation the applier gave a distinct suggestion id restages under
      // that id, so the record states it.
      const recorded =
        outcome.suggestionId === (operation.suggestionId ?? operation.id)
          ? operation
          : { ...operation, suggestionId: outcome.suggestionId };
      const record: FolioPendingSuggestionRecord = {
        version: FOLIO_PENDING_SUGGESTION_VERSION,
        suggestionId: outcome.suggestionId,
        operation: structuredClone(recorded),
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
        ...(outcome.commentId !== undefined && {
          commentId: outcome.commentId,
        }),
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
      if (
        !activeSuggestionIds.has(record.suggestionId) ||
        record.commentId === undefined
      )
        continue;
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
        isObject(value) && typeof value["suggestionId"] === "string"
          ? value["suggestionId"]
          : null;
      if (
        !isObject(value) ||
        value["version"] !== FOLIO_PENDING_SUGGESTION_VERSION
      ) {
        return { status: "stale", suggestionId, reason: "unsupportedVersion" };
      }
      const record = parseRecord(value);
      if (!record)
        return { status: "stale", suggestionId, reason: "invalidRecord" };
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
      if (!snapshot)
        return { status: "stale", suggestionId, reason: "missingAnchor" };
      const multiplicity =
        paraIdCountsOf(snapshot).get(record.anchor.paraId) ?? 0;
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
        hashFolioAIBlockText(anchor.text.slice(startOffset, endOffset)) !==
          selectedTextHash
      ) {
        return { status: "stale", suggestionId, reason: "textChanged" };
      }
      if (baseline.fingerprint !== record.sourceDocumentFingerprint) {
        return { status: "stale", suggestionId, reason: "documentChanged" };
      }
      const outcome = apply(record, snapshot);
      const applied = outcome.applied.find(
        (entry) => entry.id === record.operation.id,
      );
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
            ...(applied.commentId !== undefined && {
              commentId: applied.commentId,
            }),
          }),
        );
      }
      changedStories.add(storyKey);
      return { status: "restaged", suggestionId: record.suggestionId };
    });
  }
}
