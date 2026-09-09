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

import type { RunPropertyChange } from "../types/document";
import { expectRunPropertyChangeMarkAttrs } from "../prosemirror/attrs";
import {
  expandRunFormattingCarrier,
  runFormattingCarrierReviewText,
  runFormattingInlineAtomDisposition,
} from "../prosemirror/runFormattingInlineCarriers";
import {
  getFolioNodeRevisionCarriers,
  type FolioNodeRevisionKind,
} from "../prosemirror/revisionCarriers";
import { getTableCellMergeChange } from "../prosemirror/tableCellMergeRevision";
import { createFolioAIEditSnapshot } from "./snapshot";

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
export const FOLIO_REVIEW_CHANGE_KINDS = Object.freeze({
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
  const starts = new Map<number, string>();
  for (const anchor of Object.values(createFolioAIEditSnapshot(doc).anchors)) {
    starts.set(anchor.from, anchor.id);
  }
  return starts;
};

type FirstBlockIdWithinParams = {
  node: PMNode;
  nodePos: number;
  blockStarts: ReadonlyMap<number, string>;
};

const firstBlockIdWithin = ({
  node,
  nodePos,
  blockStarts,
}: FirstBlockIdWithinParams): string | null => {
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
};

/**
 * Whether a run's revision inside a marked row is that row's own. Folio writes
 * both halves under one revision id; Word mints a fresh `w:id` per element, so
 * an author + timestamp match counts too. A third party's edit inside the row
 * matches neither and keeps its own entry.
 */
const belongsToRowRevision = (
  scope: RowRevisionScope,
  revision: { id: number; author: string; date: string | null },
): boolean =>
  revision.id === scope.change.id ||
  (revision.author === scope.change.author && revision.date === scope.change.date);

/**
 * The tracked changes present in the body, read from inline marks and
 * structural node attributes. Runs of one inline revision within a block fold
 * into a single entry.
 *
 * A tracked row insertion or deletion marks the row AND every run in its
 * cells, the way Word writes it. Both halves are one change, so the run marks
 * fold into the row's entry rather than reporting a second time.
 */
export const getTrackedChangesFromDoc = (doc: PMNode): FolioReviewChange[] => {
  const insertionType = doc.type.schema.marks["insertion"];
  const deletionType = doc.type.schema.marks["deletion"];
  const blockStarts = blockStartIdsFromDoc(doc);
  const grouped = new Map<string, FolioReviewChange>();
  const structuralChangeCellPositions = new Map<string, Set<number>>();
  // Innermost enclosing marked row first: a table nested in a marked row can
  // carry a marked row of its own.
  const rowRevisionScopes: RowRevisionScope[] = [];
  const derivedFormattingRepresentationPositions = new Set<number>();
  let currentBlockId: string | null = null;

  doc.descendants((node, pos) => {
    while (pos >= (rowRevisionScopes.at(-1)?.end ?? Number.POSITIVE_INFINITY)) {
      rowRevisionScopes.pop();
    }
    const nodeRevisionCarriers = getFolioNodeRevisionCarriers(node, pos);
    if (nodeRevisionCarriers.length > 0) {
      const blockId = node.isTextblock
        ? (blockStarts.get(pos) ?? null)
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
        grouped.set(`row:${kind}:${revisionId}:${String(pos)}`, change);
        rowRevisionScopes.push({
          end: pos + node.nodeSize,
          kind: kind === "rowInserted" ? "insertion" : "deletion",
          change,
        });
      }
    }
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      const marker = node.attrs["cellMarker"];
      if (
        typeof marker === "object" &&
        marker !== null &&
        "kind" in marker &&
        (marker.kind === "ins" || marker.kind === "del" || marker.kind === "merge") &&
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
        const cellPositions = structuralChangeCellPositions.get(key) ?? new Set<number>();
        const shouldAppendText = existing !== undefined && !cellPositions.has(pos);
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
          const cellPositions = structuralChangeCellPositions.get(key) ?? new Set<number>();
          const shouldAppendText = existing !== undefined && !cellPositions.has(pos);
          grouped.set(key, {
            id: revisionId,
            type: "cellMerged",
            author: change.info.author,
            date: change.info.date ?? null,
            text:
              shouldAppendText && existing.text.length > 0 && text.length > 0
                ? `${existing.text}\n${text}`
                : existing?.text || text,
            blockId: existing?.blockId ?? firstBlockIdWithin({ node, nodePos: pos, blockStarts }),
          });
          cellPositions.add(pos);
          structuralChangeCellPositions.set(key, cellPositions);
        }
      }
    }
    if (node.isTextblock) {
      currentBlockId = blockStarts.get(pos) ?? null;
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
            derivedFormattingRepresentationPositions.add(representation.position);
          }
          const changeMark = representation.node.marks.find(
            (mark) => mark.type.name === "runPropertyChange",
          );
          if (!changeMark) {
            continue;
          }
          for (const change of expectRunPropertyChangeMarkAttrs(changeMark).changes) {
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

    const text = node.text ?? "";
    for (const mark of node.marks) {
      if (mark.type.name === "runPropertyChange") {
        continue;
      }
      if (typeof mark.attrs["revisionId"] !== "number") {
        continue;
      }
      let kind: FolioReviewChangeKind;
      if (mark.type === insertionType) {
        kind = "insertion";
      } else if (mark.type === deletionType) {
        kind = "deletion";
      } else {
        continue;
      }
      const revisionId = mark.attrs["revisionId"];
      const author = mark.attrs["author"];
      const date = mark.attrs["date"];
      const revision = {
        id: revisionId,
        author: typeof author === "string" ? author : "",
        date: typeof date === "string" ? date : null,
      };
      const rowScope = rowRevisionScopes.at(-1);
      if (rowScope?.kind === kind && belongsToRowRevision(rowScope, revision)) {
        continue;
      }
      const key = `${currentBlockId ?? ""}:${kind}:${revisionId}`;
      const existing = grouped.get(key);
      if (existing) {
        existing.text += text;
        continue;
      }
      grouped.set(key, { ...revision, type: kind, text, blockId: currentBlockId });
    }
    return undefined;
  });

  return [...grouped.values()];
};

/**
 * The comment anchors present in the body, read from the `comment` mark the
 * editor renders. Each entry carries the anchored text and containing block id;
 * runs of one comment id within the body fold into a single anchor.
 */
export const getCommentAnchorsFromDoc = (doc: PMNode): FolioCommentAnchor[] => {
  const commentType = doc.type.schema.marks["comment"];
  if (!commentType) {
    return [];
  }
  const blockStarts = blockStartIdsFromDoc(doc);
  const anchors = new Map<number, FolioCommentAnchor>();
  let currentBlockId: string | null = null;

  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      currentBlockId = blockStarts.get(pos) ?? null;
      return true;
    }
    if (!node.isInline || node.text === undefined) {
      return undefined;
    }
    const text = node.text;
    for (const mark of node.marks) {
      if (mark.type !== commentType || typeof mark.attrs["commentId"] !== "number") {
        continue;
      }
      const commentId = mark.attrs["commentId"];
      const existing = anchors.get(commentId);
      if (!existing) {
        anchors.set(commentId, { commentId, blockId: currentBlockId, quote: text });
        continue;
      }
      existing.quote += text;
      existing.blockId ??= currentBlockId;
    }
    return undefined;
  });

  return [...anchors.values()];
};
