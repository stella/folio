/**
 * Pure document-reading helpers over a ProseMirror `doc`. These walk a
 * `PMNode` with no `EditorView`, no DOM, and no reviewer instance, so the
 * headless {@link FolioDocxReviewer} and a live React/Vue editor can read the
 * same tracked changes and comment anchors from whichever `doc` they hold.
 *
 * Block ids are resolved through {@link createFolioAIEditSnapshot} so a change
 * or anchor reports the same stable `w14:paraId` / `seq-NNNN` block id the
 * snapshot and apply layers use.
 */

import type { Node as PMNode } from "prosemirror-model";

import {
  expectRunPropertyChangeMarkAttrs,
  expectTrackedChangeMarkAttrs,
} from "../prosemirror/attrs";
import {
  type FolioNodeRevisionKind,
  getFolioNodeRevisionCarriers,
} from "../prosemirror/revisionCarriers";
import {
  expandRunFormattingCarrier,
  runFormattingCarrierReviewText,
  runFormattingInlineAtomDisposition,
  runFormattingInlineAtomResultText,
} from "../prosemirror/runFormattingInlineCarriers";
import { getTableCellMergeChange } from "../prosemirror/tableCellMergeRevision";
import {
  enclosingInsertionAncestors,
  trackedRevisionLayerOf,
} from "../prosemirror/trackedRevisionPath";
import type { RunPropertyChange } from "../types/document";
import { commentAnchorIdsForNode } from "./comment-lifecycle";
import { readerTextOf } from "./note-references";
import {
  collectNoteReferenceLabels,
  createFolioAIEditSnapshot,
  sourceDocumentOf,
} from "./snapshot";
import type { FolioAIEditSnapshot } from "./types";

export type FolioReviewChangeKind =
  | "insertion"
  | "deletion"
  | "formatting"
  | "rowInserted"
  | "rowDeleted"
  | "cellInserted"
  | "cellDeleted"
  | "cellMerged"
  | FolioNodeRevisionKind;

/**
 * Runtime census of every tracked-change discriminator the reviewer exposes.
 *
 * The total Record is intentional: adding a new FolioReviewChangeKind without
 * adding it here is a compile-time error. Cross-product persistence tests use
 * this census, so a newly modeled revision cannot silently miss the resolver
 * and serialization matrix.
 */
export const FOLIO_REVIEW_CHANGE_KINDS: Readonly<{
  readonly insertion: "insertion";
  readonly deletion: "deletion";
  readonly formatting: "formatting";
  readonly rowInserted: "rowInserted";
  readonly rowDeleted: "rowDeleted";
  readonly cellInserted: "cellInserted";
  readonly cellDeleted: "cellDeleted";
  readonly cellMerged: "cellMerged";
  readonly paragraphMarkInserted: "paragraphMarkInserted";
  readonly paragraphMarkDeleted: "paragraphMarkDeleted";
  readonly paragraphPropertiesChanged: "paragraphPropertiesChanged";
  readonly sectionPropertiesChanged: "sectionPropertiesChanged";
  readonly tablePropertiesChanged: "tablePropertiesChanged";
  readonly tablePropertyExceptionsChanged: "tablePropertyExceptionsChanged";
  readonly rowPropertiesChanged: "rowPropertiesChanged";
  readonly cellPropertiesChanged: "cellPropertiesChanged";
}> = Object.freeze({
  insertion: "insertion",
  deletion: "deletion",
  formatting: "formatting",
  rowInserted: "rowInserted",
  rowDeleted: "rowDeleted",
  cellInserted: "cellInserted",
  cellDeleted: "cellDeleted",
  cellMerged: "cellMerged",
  paragraphMarkInserted: "paragraphMarkInserted",
  paragraphMarkDeleted: "paragraphMarkDeleted",
  paragraphPropertiesChanged: "paragraphPropertiesChanged",
  sectionPropertiesChanged: "sectionPropertiesChanged",
  tablePropertiesChanged: "tablePropertiesChanged",
  tablePropertyExceptionsChanged: "tablePropertyExceptionsChanged",
  rowPropertiesChanged: "rowPropertiesChanged",
  cellPropertiesChanged: "cellPropertiesChanged",
} as const satisfies Record<FolioReviewChangeKind, FolioReviewChangeKind>);

/** A tracked change discovered in the document body. */
export type FolioReviewChange = {
  /**
   * The tracked-change revision id — the OOXML `w:id` carried on the
   * revision marker. A text replacement produces two changes (a deletion side
   * and an insertion side) with distinct ids.
   */
  id: number;
  type: FolioReviewChangeKind;
  author: string;
  /** ISO date the change was authored, or `null` when the source omitted it. */
  date: string | null;
  /** Affected text, including row or cell content for structural revisions. */
  text: string;
  /**
   * Stable id of the containing body block (Word `w14:paraId` or `seq-NNNN`).
   * `null` when the change sits in a block with no surviving visible text,
   * which has no snapshot block.
   */
  blockId: string | null;
};

/** A tracked change an accept-all or reject-all left pending, and why. */
export type FolioReviewResolutionSkip = {
  /** The change's revision id, as {@link FolioReviewChange.id}. */
  id: number;
  /**
   * `contentControlLock`: resolving it would change the content of a content
   * control whose `w:lock` is `contentLocked` or `sdtContentLocked`, or remove
   * one whose lock is `sdtLocked` or `sdtContentLocked`.
   */
  reason: "contentControlLock";
  message: string;
};

/** What an accept-all or reject-all did. */
export type FolioReviewResolution = {
  /** The ordinary tracked changes it resolved. */
  resolved: number;
  /** The changes it left pending, each as it was. */
  skipped: FolioReviewResolutionSkip[];
};

/** A comment anchor (the ranged text a comment marks) discovered in the body. */
export type FolioCommentAnchor = {
  commentId: number;
  /** Stable id of the anchored body block, or `null` when the anchor is absent. */
  blockId: string | null;
  /** The document text the comment is anchored to. */
  quote: string;
};

/** Map each snapshot block's start position to its stable id. */
const blockStartIdsFromDoc = (doc: PMNode): Map<number, string> => {
  return blockStartIdsFromSnapshot(createFolioAIEditSnapshot(doc));
};

/** Map an existing snapshot's block starts without projecting its document again. */
const blockStartIdsFromSnapshot = (
  snapshot: FolioAIEditSnapshot,
): Map<number, string> => {
  const starts = new Map<number, string>();
  for (const anchor of Object.values(snapshot.anchors)) {
    starts.set(anchor.from, anchor.id);
  }
  return starts;
};

type FirstBlockIdWithinParams = {
  node: PMNode;
  nodePos: number;
  blockStarts: ReadonlyMap<number, string> | null;
};

const firstBlockIdWithin = ({
  node,
  nodePos,
  blockStarts,
}: FirstBlockIdWithinParams): string | null => {
  if (!blockStarts) {
    return null;
  }
  let blockId: string | null = null;
  node.descendants((child, relativePos) => {
    if (blockId !== null || !child.isTextblock) {
      return blockId === null;
    }
    blockId = blockStarts.get(nodePos + 1 + relativePos) ?? null;
    return false;
  });
  return blockId;
};

/** An enclosing table row whose structural revision also marked its runs. */
type RowRevisionScope = {
  /** Document position one past the row, where the scope ends. */
  end: number;
  /** The run revision the row's marker wrote (`insertion` for `w:trPr/w:ins`). */
  kind: "insertion" | "deletion";
  /** The row's own change, which the folded run marks belong to. */
  change: FolioReviewChange;
  /** The row change's key among the grouped changes. */
  key: string;
};

/**
 * Whether a run's revision inside a marked row is that row's own. Folio writes
 * both halves under one revision id; Word mints a fresh `w:id` per element, so
 * an author + timestamp match counts too. A third party's edit inside the row
 * matches neither and keeps its own entry.
 */
type RevisionInfo = { id: number; author: string; date: string | null };

type InlineBoundary = {
  hyperlink: string | null;
  /**
   * Every comment the run sits in, in id order: comments nest, and a
   * comment that starts or ends inside another is a boundary too.
   */
  comments: string;
  revisionParents: string;
};

type InlineChangeRange = {
  start: number;
  end: number;
  firstBoundary: InlineBoundary;
  lastBoundary: InlineBoundary;
};

type TrackedChangeGroup = { change: FolioReviewChange; ids: number[] };

const belongsToRowRevision = (
  scope: RowRevisionScope,
  revision: RevisionInfo,
): boolean =>
  revision.id === scope.change.id ||
  (revision.author === scope.change.author &&
    revision.date === scope.change.date);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

/**
 * Every revision id a model value holds: each tracked record carries its
 * `info`, and a table grid revision its own `id`.
 */
const revisionIdsIn = (value: unknown, ids = new Set<number>()) => {
  if (Array.isArray(value)) {
    for (const item of value) revisionIdsIn(item, ids);
    return ids;
  }
  if (!isRecord(value)) return ids;
  for (const [key, field] of Object.entries(value)) {
    if (key === "sourceXml" || key === "gridSourceXml") continue;
    if (
      (key === "info" || key === "gridChange") &&
      isRecord(field) &&
      typeof field["id"] === "number"
    )
      ids.add(field["id"]);
    revisionIdsIn(field, ids);
  }
  return ids;
};

/** The grid revision a table node keeps with its source formatting, if any. */
const tableGridRevisionId = (node: PMNode): number | null => {
  const formatting: unknown = node.attrs["_originalFormatting"];
  if (!isRecord(formatting)) return null;
  const gridChange = formatting["gridChange"];
  return isRecord(gridChange) && typeof gridChange["id"] === "number"
    ? gridChange["id"]
    : null;
};

/**
 * Revision ids in the vertically merged continuation cells each row of
 * `table` holds without a merge revision of their own. The editor folds a
 * continuation cell into the cell that starts the merge, so its content has
 * no node; its row's change stands for it.
 */
const continuationRevisionIdsByRow = (
  table: PMNode,
  tablePos: number,
): Map<number, Set<number>> => {
  const rowPositions: number[] = [];
  table.forEach((_row, offset) => rowPositions.push(tablePos + 1 + offset));
  const byRow = new Map<number, Set<number>>();
  table.forEach((row, _offset, rowIndex) => {
    row.forEach((cell) => {
      const continuations: unknown = cell.attrs["_docxVMergeContinuationCells"];
      if (!Array.isArray(continuations)) return;
      continuations.forEach((continuation: unknown, index) => {
        if (getTableCellMergeChange(continuation)) return;
        const rowPos = rowPositions[rowIndex + 1 + index];
        if (rowPos === undefined) return;
        const ids = byRow.get(rowPos) ?? new Set<number>();
        revisionIdsIn(continuation, ids);
        if (ids.size > 0) byRow.set(rowPos, ids);
      });
    });
  });
  return byRow;
};

/** Shared revision interpreter; a null block map omits unrelated block-id projection. */
const getTrackedChangeGroupsFromProjectedDoc = (
  doc: PMNode,
  blockStarts: ReadonlyMap<number, string> | null,
): TrackedChangeGroup[] => {
  const insertionType = doc.type.schema.marks["insertion"];
  const deletionType = doc.type.schema.marks["deletion"];
  const grouped = new Map<string, FolioReviewChange>();
  const inlineRanges = new Map<string, InlineChangeRange>();
  const lastInlineGroupByPath = new Map<string, string>();
  const transparentInlineEnds = new Map<number, number>();
  const transparentGap = (start: number, end: number): boolean => {
    let pos = start;
    while (pos < end) {
      const next = transparentInlineEnds.get(pos);
      if (next === undefined) return false;
      pos = next;
    }
    return pos === end;
  };
  const structuralChangeCellPositions = new Map<string, Set<number>>();
  // Innermost enclosing marked row first: a table nested in a marked row can
  // carry a marked row of its own.
  const rowRevisionScopes: RowRevisionScope[] = [];
  // Revision ids an entry stands for beyond its own: the run marks a row
  // folds, the content of continuation cells, a table's grid revision.
  // Resolving the entry reaches them.
  const foldedIds = new Map<string, Set<number>>();
  const fold = (key: string, ids: Iterable<number>) => {
    const folded = foldedIds.get(key) ?? new Set<number>();
    for (const id of ids) folded.add(id);
    foldedIds.set(key, folded);
  };
  const continuationRevisions = new Map<number, Set<number>>();
  const derivedFormattingRepresentationPositions = new Set<number>();
  const noteReferences = collectNoteReferenceLabels(doc);
  let currentBlockId: string | null = null;

  doc.descendants((node, pos) => {
    if (node.type.name === "commentReference") {
      transparentInlineEnds.set(pos, pos + node.nodeSize);
    }
    while (pos >= (rowRevisionScopes.at(-1)?.end ?? Number.POSITIVE_INFINITY)) {
      rowRevisionScopes.pop();
    }
    const nodeRevisionCarriers = getFolioNodeRevisionCarriers(node, pos);
    if (nodeRevisionCarriers.length > 0) {
      const blockId = node.isTextblock
        ? (blockStarts?.get(pos) ?? null)
        : firstBlockIdWithin({ node, nodePos: pos, blockStarts });
      nodeRevisionCarriers.forEach((carrier, carrierIndex) => {
        grouped.set(
          `node:${String(pos)}:${carrier.type}:${String(carrier.id)}:${String(carrierIndex)}`,
          {
            id: carrier.id,
            type: carrier.type,
            author: carrier.author,
            date: carrier.date,
            text: carrier.text,
            blockId,
          },
        );
      });
    }
    if (node.type.name === "table") {
      for (const [rowPos, ids] of continuationRevisionIdsByRow(node, pos))
        continuationRevisions.set(rowPos, ids);
      const gridId = tableGridRevisionId(node);
      if (gridId !== null) {
        // A grid revision has no author or date of its own; it is part of
        // the table's property change, or lists as one when there is none.
        const tableChangeKey = [...grouped.entries()].find(
          ([key, change]) =>
            key.startsWith(`node:${String(pos)}:`) &&
            change.type === "tablePropertiesChanged",
        )?.[0];
        if (tableChangeKey !== undefined) fold(tableChangeKey, [gridId]);
        else
          grouped.set(`grid:${String(gridId)}:${String(pos)}`, {
            id: gridId,
            type: "tablePropertiesChanged",
            author: "",
            date: null,
            text: node.textContent,
            blockId: firstBlockIdWithin({ node, nodePos: pos, blockStarts }),
          });
      }
    }
    if (node.type.name === "tableRow") {
      for (const [attrName, kind] of [
        ["trIns", "rowInserted"],
        ["trDel", "rowDeleted"],
      ] as const) {
        const marker = node.attrs[attrName];
        if (
          typeof marker !== "object" ||
          marker === null ||
          !("revisionId" in marker) ||
          typeof marker.revisionId !== "number"
        ) {
          continue;
        }
        const revisionId = marker.revisionId;
        const author = "author" in marker ? marker.author : undefined;
        const date = "date" in marker ? marker.date : undefined;
        const change: FolioReviewChange = {
          id: revisionId,
          type: kind,
          author: typeof author === "string" ? author : "",
          date: typeof date === "string" ? date : null,
          text: node.textContent,
          blockId: firstBlockIdWithin({ node, nodePos: pos, blockStarts }),
        };
        const key = `row:${kind}:${revisionId}:${String(pos)}`;
        grouped.set(key, change);
        fold(key, continuationRevisions.get(pos) ?? []);
        rowRevisionScopes.push({
          end: pos + node.nodeSize,
          kind: kind === "rowInserted" ? "insertion" : "deletion",
          change,
          key,
        });
      }
    }
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      const marker = node.attrs["cellMarker"];
      if (
        typeof marker === "object" &&
        marker !== null &&
        "kind" in marker &&
        (marker.kind === "ins" ||
          marker.kind === "del" ||
          marker.kind === "merge") &&
        "info" in marker &&
        typeof marker.info === "object" &&
        marker.info !== null &&
        "revisionId" in marker.info &&
        typeof marker.info.revisionId === "number"
      ) {
        const revisionId = marker.info.revisionId;
        const author = "author" in marker.info ? marker.info.author : undefined;
        const date = "date" in marker.info ? marker.info.date : undefined;
        let kind: FolioReviewChangeKind = "cellMerged";
        if (marker.kind === "ins") {
          kind = "cellInserted";
        } else if (marker.kind === "del") {
          kind = "cellDeleted";
        }
        const key = `cell:${kind}:${revisionId}`;
        const text = node.textContent;
        const blockId = firstBlockIdWithin({ node, nodePos: pos, blockStarts });
        const existing = grouped.get(key);
        const cellPositions =
          structuralChangeCellPositions.get(key) ?? new Set<number>();
        const shouldAppendText =
          existing !== undefined && !cellPositions.has(pos);
        grouped.set(key, {
          id: revisionId,
          type: kind,
          author: typeof author === "string" ? author : "",
          date: typeof date === "string" ? date : null,
          text:
            shouldAppendText && existing.text.length > 0 && text.length > 0
              ? `${existing.text}\n${text}`
              : existing?.text || text,
          blockId: existing?.blockId ?? blockId,
        });
        cellPositions.add(pos);
        structuralChangeCellPositions.set(key, cellPositions);
      }
      const continuationCells = node.attrs["_docxVMergeContinuationCells"];
      if (Array.isArray(continuationCells)) {
        for (const continuationCell of continuationCells) {
          const change = getTableCellMergeChange(continuationCell);
          if (!change) {
            continue;
          }
          const revisionId = change.info.id;
          const key = `cell:cellMerged:${revisionId}`;
          const existing = grouped.get(key);
          const text = node.textContent;
          const cellPositions =
            structuralChangeCellPositions.get(key) ?? new Set<number>();
          const shouldAppendText =
            existing !== undefined && !cellPositions.has(pos);
          grouped.set(key, {
            id: revisionId,
            type: "cellMerged",
            author: change.info.author,
            date: change.info.date ?? null,
            text:
              shouldAppendText && existing.text.length > 0 && text.length > 0
                ? `${existing.text}\n${text}`
                : existing?.text || text,
            blockId:
              existing?.blockId ??
              firstBlockIdWithin({ node, nodePos: pos, blockStarts }),
          });
          cellPositions.add(pos);
          structuralChangeCellPositions.set(key, cellPositions);
          // The continuation cell has no node of its own: what else it
          // holds belongs to the merge that folded it.
          fold(key, revisionIdsIn(continuationCell));
        }
      }
    }
    if (node.isTextblock) {
      currentBlockId = blockStarts?.get(pos) ?? null;
      return true;
    }
    if (!node.isInline) {
      return undefined;
    }
    if (!derivedFormattingRepresentationPositions.has(pos)) {
      const disposition = runFormattingInlineAtomDisposition(node);
      const formattingCarrier =
        disposition === "structured-field" ||
        (disposition !== null &&
          disposition !== "not-a-run" &&
          node.marks.some((mark) => mark.type.name === "runPropertyChange"))
          ? expandRunFormattingCarrier(node, pos)
          : null;
      if (formattingCarrier) {
        let changesById: Map<number, RunPropertyChange> | undefined;
        for (const representation of formattingCarrier.representations) {
          if (representation.role === "serialized-result") {
            derivedFormattingRepresentationPositions.add(
              representation.position,
            );
          }
          const changeMark = representation.node.marks.find(
            (mark) => mark.type.name === "runPropertyChange",
          );
          if (!changeMark) {
            continue;
          }
          for (const change of expectRunPropertyChangeMarkAttrs(changeMark)
            .changes) {
            changesById ??= new Map();
            if (!changesById.has(change.info.id)) {
              changesById.set(change.info.id, change);
            }
          }
        }
        if (changesById) {
          const text = runFormattingCarrierReviewText(formattingCarrier);
          for (const change of changesById.values()) {
            const key = `${currentBlockId ?? ""}:formatting:${String(change.info.id)}`;
            const existing = grouped.get(key);
            if (existing) {
              existing.text += text;
              continue;
            }
            grouped.set(key, {
              id: change.info.id,
              type: "formatting",
              author: change.info.author,
              date: change.info.date ?? null,
              text,
              blockId: currentBlockId,
            });
          }
        }
      }
    }

    const text = readerTextOf(node, noteReferences);
    const insertionMark = node.marks.find(
      (mark) => mark.type === insertionType,
    );
    const deletionMark = node.marks.find((mark) => mark.type === deletionType);
    const mark = deletionMark ?? insertionMark;
    if (!mark || typeof mark.attrs["revisionId"] !== "number") {
      return undefined;
    }
    const attrs = expectTrackedChangeMarkAttrs(mark);
    // Live edits carry two marks; parsed nested revisions carry one mark
    // and its ancestor path. Use the serializer's path in both forms.
    const pathLayers = [
      ...enclosingInsertionAncestors({
        insertionMark: deletionMark ? insertionMark : undefined,
        ancestors: attrs._docxRevisionAncestors ?? [],
        node,
      }),
      trackedRevisionLayerOf(mark, node),
    ];
    const layers = pathLayers.map((layer) => ({
      kind:
        layer.type === "insertion" || layer.type === "moveTo"
          ? ("insertion" as const)
          : ("deletion" as const),
      revision: {
        id: layer.revisionId,
        author: layer.author,
        date: layer.date ?? null,
      },
    }));
    for (const [
      layerIndex,
      { kind: layerKind, revision },
    ] of layers.entries()) {
      const rowScope = rowRevisionScopes.at(-1);
      if (
        rowScope?.kind === layerKind &&
        belongsToRowRevision(rowScope, revision)
      ) {
        // The row's entry stands for this run's revision too, so its ids
        // reach a run mark minted under another id.
        fold(rowScope.key, [revision.id]);
        continue;
      }
      const path = `${currentBlockId ?? ""}:${JSON.stringify(pathLayers.slice(0, layerIndex + 1))}`;
      const previousKey = lastInlineGroupByPath.get(path);
      const previousRange =
        previousKey === undefined ? undefined : inlineRanges.get(previousKey);
      // A wrapper closes when its ancestor path changes or unmarked text
      // intervenes. Reusing its logical id later does not reopen that wrapper.
      const key =
        previousKey !== undefined &&
        previousRange &&
        transparentGap(previousRange.end, pos)
          ? previousKey
          : `${path}:${pos}`;
      lastInlineGroupByPath.set(path, key);
      const hyperlink = node.marks.find(
        (candidate) => candidate.type.name === "hyperlink",
      );
      const commentIds = node.marks.flatMap((candidate) => {
        const commentId = candidate.attrs["commentId"];
        return candidate.type.name === "comment" &&
          typeof commentId === "number"
          ? [commentId]
          : [];
      });
      const boundary: InlineBoundary = {
        hyperlink: hyperlink ? String(hyperlink.attrs["href"] ?? "") : null,
        comments: commentIds.toSorted((left, right) => left - right).join(","),
        revisionParents: JSON.stringify(pathLayers.slice(0, layerIndex)),
      };
      const range = inlineRanges.get(key);
      if (range) {
        range.end = pos + node.nodeSize;
        range.lastBoundary = boundary;
      } else {
        inlineRanges.set(key, {
          start: pos,
          end: pos + node.nodeSize,
          firstBoundary: boundary,
          lastBoundary: boundary,
        });
      }
      const existing = grouped.get(key);
      if (existing) {
        existing.text += text;
        continue;
      }
      grouped.set(key, {
        ...revision,
        type: layerKind,
        text,
        blockId: currentBlockId,
      });
    }
    return undefined;
  });

  // OOXML assigns distinct physical ids when an inline revision crosses a
  // hyperlink or comment boundary. Keep its contiguous fragments actionable
  // as one logical change, without joining separate adjacent revisions.
  const groups: TrackedChangeGroup[] = [];
  let previousInline: { key: string; group: TrackedChangeGroup } | null = null;
  for (const [key, change] of grouped) {
    const range = inlineRanges.get(key);
    // Formatting is an independent revision stream: it cannot interrupt the
    // contiguous inline wrapper whose run carries the property revision.
    const previous = previousInline?.group;
    const previousRange =
      previousInline === null
        ? undefined
        : inlineRanges.get(previousInline.key);
    if (
      previous &&
      range &&
      previousRange &&
      previousRange.end <= range.start &&
      transparentGap(previousRange.end, range.start) &&
      previousRange.lastBoundary.revisionParents ===
        range.firstBoundary.revisionParents &&
      (previousRange.lastBoundary.hyperlink !== range.firstBoundary.hyperlink ||
        previousRange.lastBoundary.comments !== range.firstBoundary.comments ||
        previousRange.end !== range.start) &&
      (change.type === "insertion" || change.type === "deletion") &&
      previous.change.type === change.type &&
      previous.change.author === change.author &&
      previous.change.date === change.date &&
      previous.change.blockId === change.blockId
    ) {
      previous.change.text += change.text;
      previous.ids.push(change.id, ...(foldedIds.get(key) ?? []));
      previousRange.end = range.end;
      previousRange.lastBoundary = range.lastBoundary;
      continue;
    }
    const group = {
      change: { ...change },
      ids: [
        change.id,
        ...[...(foldedIds.get(key) ?? [])].filter((id) => id !== change.id),
      ],
    };
    groups.push(group);
    if (range) {
      previousInline = { key, group };
    }
  }
  return groups;
};

const getTrackedChangesFromProjectedDoc = (
  doc: PMNode,
  blockStarts: ReadonlyMap<number, string> | null,
): FolioReviewChange[] =>
  getTrackedChangeGroupsFromProjectedDoc(doc, blockStarts).map(
    ({ change }) => change,
  );

/** Read revisions from the exact immutable document that produced a snapshot. */
export const getTrackedChangesFromSnapshot = (
  snapshot: FolioAIEditSnapshot,
): FolioReviewChange[] =>
  getTrackedChangesFromProjectedDoc(
    sourceDocumentOf(snapshot),
    blockStartIdsFromSnapshot(snapshot),
  );

/**
 * The tracked changes present in the body, read from inline marks and
 * structural node attributes. Adjacent runs sharing a revision wrapper path
 * fold into one entry; interrupted or differently nested wrappers stay separate.
 *
 * A tracked row insertion or deletion marks the row AND every run in its
 * cells, the way Word writes it. Both halves are one change, so the run marks
 * fold into the row's entry rather than reporting a second time.
 */
export const getTrackedChangesFromDoc = (doc: PMNode): FolioReviewChange[] =>
  getTrackedChangesFromProjectedDoc(doc, blockStartIdsFromDoc(doc));

/**
 * The physical revision ids behind each listed change that `ids` name, one
 * group per change. An inline change split by markup is one entry of the
 * change list over several ids, as is a marked row with the run marks it
 * folds, and any of them resolves the whole change. An id no change lists
 * stands alone.
 */
export const getTrackedChangeGroupsForIds = (
  doc: PMNode,
  ids: readonly number[],
): number[][] => {
  const groups = getTrackedChangeGroupsFromProjectedDoc(
    doc,
    blockStartIdsFromDoc(doc),
  );
  const out: number[][] = [];
  const taken = new Set<number>();
  for (const id of ids) {
    if (taken.has(id)) continue;
    const members = groups.find((group) => group.ids.includes(id))?.ids ?? [id];
    const fresh = [...new Set(members)].filter((member) => !taken.has(member));
    for (const member of fresh) taken.add(member);
    out.push(fresh);
  }
  return out;
};

/** Physical revision ids behind an inline change split by markup. */
export const getTrackedChangeGroupIdsFromDoc = (
  doc: PMNode,
  id: number,
): number[] => getTrackedChangeGroupsForIds(doc, [id]).at(0) ?? [id];

/** @internal Read revision statistics without paying for an unrelated block projection. */
export const getTrackedChangeStatsFromDoc = (
  doc: PMNode,
): { highestId: number; present: boolean; ids: readonly number[] } => {
  let highestId = 0;
  const revisionIds = new Set<number>();
  const groups = getTrackedChangeGroupsFromProjectedDoc(doc, null);
  for (const { ids } of groups) {
    for (const id of ids) {
      highestId = Math.max(highestId, id);
      revisionIds.add(id);
    }
  }
  return { highestId, present: groups.length > 0, ids: [...revisionIds] };
};

/**
 * The comment anchors present in the body, read from the `comment` mark the
 * editor renders. Each entry carries the anchored text and containing block id;
 * ranges for one comment id within the body fold into a single anchor.
 */
export const getCommentAnchorsFromDoc = (doc: PMNode): FolioCommentAnchor[] => {
  const commentType = doc.type.schema.marks["comment"];
  if (!commentType) {
    return [];
  }
  const blockStarts = blockStartIdsFromDoc(doc);
  const ranges = new Map<
    number,
    { blockId: string | null; start: number; end: number }
  >();
  const noteReferences = collectNoteReferenceLabels(doc);
  let currentBlockId: string | null = null;
  let bodyText = "";

  // OOXML stores one continuous range per comment id, including unmarked
  // tracked runs and paragraphs between its first and last marked text.
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      currentBlockId = blockStarts.get(pos) ?? null;
      return true;
    }
    for (const commentId of commentAnchorIdsForNode(node)) {
      if (!ranges.has(commentId)) {
        ranges.set(commentId, {
          blockId: currentBlockId,
          start: bodyText.length,
          end: bodyText.length,
        });
      }
    }
    // A field reads as its result. One with its result in runs adds them as its
    // children, but a comment over the field marks the field, not them.
    const isStructuredField = node.type.name === "structuredField";
    const fieldText = isStructuredField
      ? node.textContent
      : runFormattingInlineAtomResultText(node);
    if (!node.isInline || (node.text === undefined && fieldText === null)) {
      return undefined;
    }
    const text =
      node.text === undefined
        ? (fieldText ?? "")
        : readerTextOf(node, noteReferences);
    const end = bodyText.length + text.length;
    for (const mark of node.marks) {
      if (
        mark.type !== commentType ||
        typeof mark.attrs["commentId"] !== "number"
      ) {
        continue;
      }
      const commentId = mark.attrs["commentId"];
      const range = ranges.get(commentId);
      if (range) {
        range.end = end;
        range.blockId ??= currentBlockId;
      } else {
        ranges.set(commentId, {
          blockId: currentBlockId,
          start: bodyText.length,
          end,
        });
      }
    }
    if (!isStructuredField) bodyText += text;
    return undefined;
  });

  return [...ranges].map(([commentId, { blockId, start, end }]) => ({
    commentId,
    blockId,
    quote: bodyText.slice(start, end),
  }));
};
