import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import {
  acceptAllChanges,
  rejectEverySuggestion,
} from "../prosemirror/commands/comments";
import {
  createDocumentNumberingPlugin,
  getStatedDocumentNumbering,
} from "../prosemirror/plugins/documentNumbering";
import { createDocumentStyleContextPlugin } from "../prosemirror/plugins/documentStyles";
import { getDocumentStyleResolver } from "../prosemirror/plugins/documentStyleState";
import {
  anchoredCommentIdsInProseDoc,
  commentAnchorIdsForNode,
  coveredCommentIdsInProseDoc,
  isCommentAnchorCarrier,
} from "./comment-lifecycle";
import { comparisonApplying } from "./note-references";
import type {
  FolioAIEditAppliedOperation,
  FolioAIEditOperation,
} from "./types";

const deletionOperations: ReadonlySet<FolioAIEditOperation["type"]> = new Set([
  "deleteBlock",
  "deleteTable",
  "deleteTableRow",
  "deleteTableColumn",
  "replaceBlock",
  "replaceInBlock",
  "replaceRange",
  "splitBlock",
]);

type Deletion = { owner: string | undefined };
// Null means coverage survives or belongs to more than one operation.
type AnchorCoverage = Map<number, string | null>;

const objectAttr = (value: unknown, key: string): unknown =>
  typeof value === "object" && value !== null
    ? Reflect.get(value, key)
    : undefined;

const deletedCell = (node: PMNode): unknown => {
  const marker: unknown = node.attrs["cellMarker"];
  return objectAttr(marker, "kind") === "del"
    ? objectAttr(marker, "info")
    : undefined;
};

const deletedParagraphMark = (node: PMNode): unknown => {
  const marker: unknown = node.attrs["pPrMark"];
  return objectAttr(marker, "kind") === "del"
    ? objectAttr(marker, "info")
    : undefined;
};

type TextSpan = { from: number; to: number };

type RetiredCommentReceiptsOptions = {
  before: EditorState;
  after: PMNode;
  applied: readonly FolioAIEditAppliedOperation[];
  operations: readonly Pick<FolioAIEditOperation, "id" | "type">[];
  /** Canonical inline changes can relocate markers before later steps run. */
  sourceRanges?: ReadonlyMap<string, readonly TextSpan[]>;
  /** Canonical review can supply its accepted package projection. */
  acceptedCommentIds?: (side: "before" | "after") => ReadonlySet<number> | null;
};

const acceptedComments = (
  before: EditorState,
  doc: PMNode,
): ReadonlySet<number> | null => {
  const state = rejectEverySuggestion(
    EditorState.create({
      doc,
      plugins: [
        createDocumentNumberingPlugin(getStatedDocumentNumbering(before)),
        createDocumentStyleContextPlugin(getDocumentStyleResolver(before)),
      ],
    }),
  );
  let accepted = state.doc;
  const resolved = acceptAllChanges()(state, (transaction) => {
    accepted = transaction.doc;
  });
  return resolved ? anchoredCommentIdsInProseDoc(accepted) : null;
};

const candidatesOf = (coverage: AnchorCoverage) =>
  [...coverage].flatMap(([commentId, owner]) =>
    owner === null ? [] : [{ commentId, owner }],
  );

/** Report full anchor removal only where one operation owns the deletion. */
export const addRetiredCommentReceipts = ({
  before,
  after,
  applied,
  operations,
  sourceRanges,
  acceptedCommentIds,
}: RetiredCommentReceiptsOptions): void => {
  if (comparisonApplying() || before.doc === after) return;
  const deletions = new Set(
    operations
      .filter(({ type }) => deletionOperations.has(type))
      .map(({ id }) => id),
  );
  const byRevision = new Map<number, string>();
  for (const { id, revisionIds } of applied) {
    if (!deletions.has(id)) continue;
    for (const revisionId of revisionIds ?? []) byRevision.set(revisionId, id);
  }
  if (byRevision.size === 0) return;
  const from = before.doc.content.findDiffStart(after.content);
  if (from === null) return;
  const changedEnd = before.doc.content.findDiffEnd(after.content);
  if (changedEnd === null) return;
  let touchedAnchor = false;
  after.nodesBetween(from, Math.max(from, changedEnd.b), (node) => {
    touchedAnchor ||=
      node.type.name === "commentReference" ||
      node.marks.some(({ type }) => type.name === "comment");
    return !touchedAnchor;
  });
  if (!touchedAnchor) return;

  const coverage: AnchorCoverage = new Map();
  const points: AnchorCoverage = new Map();
  const recordCoverage = (
    anchors: AnchorCoverage,
    commentId: number,
    deletion: Deletion | undefined,
  ) => {
    const previous = anchors.get(commentId);
    if (previous === null) return;
    if (deletion === undefined) {
      anchors.set(commentId, null);
      return;
    }
    const { owner } = deletion;
    if (owner === undefined) return;
    anchors.set(
      commentId,
      previous === undefined || previous === owner ? owner : null,
    );
  };
  const deletionOf = (revisionId: unknown): Deletion => ({
    owner:
      typeof revisionId === "number" ? byRevision.get(revisionId) : undefined,
  });
  const visit = (
    node: PMNode,
    ancestorDeletion: Deletion | undefined,
    pointDeletion: Deletion | undefined,
  ): void => {
    const revision: unknown =
      node.type.name === "tableRow" ? node.attrs["trDel"] : deletedCell(node);
    const structure =
      ancestorDeletion ??
      (revision == null
        ? undefined
        : deletionOf(objectAttr(revision, "revisionId")));
    // A paragraph-mark deletion supplies ownership for point references.
    const paragraphMark = deletedParagraphMark(node);
    const point =
      structure ??
      pointDeletion ??
      (paragraphMark === undefined
        ? undefined
        : deletionOf(objectAttr(paragraphMark, "id")));
    if (node.type.name === "commentReference") {
      for (const commentId of commentAnchorIdsForNode(node))
        recordCoverage(points, commentId, point);
    } else if (node.isInline && !isCommentAnchorCarrier(node)) {
      const mark = node.marks.find(({ type }) => type.name === "deletion");
      const deletion =
        structure ??
        (mark === undefined ? undefined : deletionOf(mark.attrs["revisionId"]));
      for (const comment of node.marks) {
        if (comment.type.name !== "comment") continue;
        const id: unknown = comment.attrs["commentId"];
        if (typeof id === "number") recordCoverage(coverage, id, deletion);
      }
    }
    node.forEach((child) => visit(child, structure, point));
  };
  visit(after, undefined, undefined);

  const candidates = candidatesOf(coverage);
  const possiblePoints = candidatesOf(points).filter(
    ({ commentId }) => !coverage.has(commentId),
  );
  if (possiblePoints.length > 0) {
    const rangedBefore = coveredCommentIdsInProseDoc(before.doc);
    before.doc.descendants((node) => {
      if (node.type.name !== "commentReference" && isCommentAnchorCarrier(node))
        for (const id of commentAnchorIdsForNode(node)) rangedBefore.add(id);
      return true;
    });
    for (const candidate of possiblePoints)
      if (!rangedBefore.has(candidate.commentId)) candidates.push(candidate);
  }
  if (candidates.length === 0) return;

  // Canonical replacement can relocate the marker onto a later revision.
  // Require that revision's operation to cover the original visible anchor.
  const originalCoverage = new Map<number, TextSpan[]>();
  if (
    sourceRanges !== undefined &&
    candidates.some(({ owner }) => sourceRanges.has(owner))
  ) {
    const inspected = new Set(candidates.map(({ commentId }) => commentId));
    before.doc.descendants((node, position) => {
      if (
        !node.isInline ||
        isCommentAnchorCarrier(node) ||
        node.marks.some(({ type }) => type.name === "deletion")
      )
        return true;
      for (const id of commentAnchorIdsForNode(node)) {
        if (!inspected.has(id)) continue;
        const span = { from: position, to: position + node.nodeSize };
        const earlier = originalCoverage.get(id);
        if (earlier === undefined) originalCoverage.set(id, [span]);
        else earlier.push(span);
      }
      return true;
    });
  }
  const ownsOriginalCoverage = (commentId: number, owner: string): boolean => {
    const ranges = sourceRanges?.get(owner);
    if (ranges === undefined) return true;
    const originalSpans = originalCoverage.get(commentId);
    if (originalSpans === undefined) return false;
    let rangeIndex = 0;
    for (const { from: start, to: end } of originalSpans) {
      let range = ranges.at(rangeIndex);
      while (range !== undefined && range.to <= start)
        range = ranges.at(++rangeIndex);
      if (range === undefined || range.from > start || range.to < end)
        return false;
    }
    return true;
  };

  const owned = candidates.filter(({ commentId, owner }) =>
    ownsOriginalCoverage(commentId, owner),
  );
  if (owned.length === 0) return;
  const originallyAccepted = acceptedCommentIds
    ? acceptedCommentIds("before")
    : acceptedComments(before, before.doc);
  const finallyAccepted = acceptedCommentIds
    ? acceptedCommentIds("after")
    : acceptedComments(before, after);
  if (originallyAccepted === null || finallyAccepted === null) return;

  const byOperation = new Map<string, number[]>();
  for (const { commentId, owner } of owned) {
    if (!originallyAccepted.has(commentId) || finallyAccepted.has(commentId))
      continue;
    const ids = byOperation.get(owner);
    if (ids === undefined) byOperation.set(owner, [commentId]);
    else ids.push(commentId);
  }
  for (const entry of applied) {
    const ids = byOperation.get(entry.id);
    if (ids !== undefined) entry.retiredCommentIds = ids.sort((a, b) => a - b);
  }
};
