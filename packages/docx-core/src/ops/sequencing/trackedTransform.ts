/**
 * Concurrent tracked changes whose two orders record different cuts.
 *
 * A tracked deletion or insertion records, on its wrapper, how many records
 * it cut at each edge and how many ran across its whole range, so that
 * resolving it joins them again. Whether a record runs across an edge depends
 * on what was applied before: another edit that cuts records at the same
 * offset (a tracked edge, an insertion point, a format's end, an inline split
 * or join, a paragraph split) leaves the cut to whichever applies first, and
 * a cut inside the range decides whether a record runs across the whole of
 * it. A direct deletion reaching an edge of a tracked change removes what
 * lay against that edge in one order and moves the edge in the other, and
 * one covering the change whole leaves the records around it cut in one
 * order only. Two tracked deletions that overlap leave the shared content
 * to whichever wraps it first: the one sequenced first.
 *
 * Such pairs are accepted and rebased like any other: the authority's order
 * decides the document, and every peer reaches it by replaying its pending
 * batches over the sequenced ones. Applied in the other order they record
 * other cuts, so a peer holding one of them pending gets the document as a
 * replacement rather than a rebased change. These are the only pairs
 * accepted without commuting.
 */
import type { Document } from "../../model/document";
import { leafSpans } from "../leaves";
import { findStoryParagraph } from "../paragraphLocator";
import { isTrackedWrapper } from "../review";
import { findStoryBody } from "../stories";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type TextPosition,
} from "../types";
import { isInlineInsertion, sameParagraph } from "./points";

/**
 * A direct erasure can remove an earlier newly tracked fragment, shifting
 * the supplied ids onto different source content when the tracked deletion
 * replays. The ordered id pool carries no source-to-fragment mapping, so a
 * pending decision cannot safely follow that replay. A nonempty revision
 * pool witnesses the allocation risk, even when some supplied ids go unused.
 */
export const fragmentedDeletionOverlap = (
  op: DocumentOp,
  over: DocumentOp,
): boolean => {
  if (
    op.type !== DOCUMENT_OP_TYPES.DELETE_RANGE ||
    over.type !== DOCUMENT_OP_TYPES.DELETE_RANGE
  )
    return false;
  const tracked = op.revision === undefined ? over : op;
  const direct = op.revision === undefined ? op : over;
  return (
    tracked.revision !== undefined &&
    direct.revision === undefined &&
    (tracked.newIds?.revision?.length ?? 0) > 0 &&
    sameParagraph(tracked.from, tracked.to) &&
    sameParagraph(tracked.from, direct.from) &&
    sameParagraph(tracked.from, direct.to) &&
    Math.max(tracked.from.offset, direct.from.offset) <
      Math.min(tracked.to.offset, direct.to.offset)
  );
};

/** The kinds of concurrent pair whose two orders record different cuts. */
export const TRACKED_ORDER_PAIRS: Readonly<{
  /** A cut at a tracked change's edge or inside its range. */
  CUT_AT_TRACKED_CHANGE: "cutAtTrackedChange";
  /** A direct deletion reaching an edge of a tracked change, or covering it. */
  DELETION_AT_TRACKED_EDGE: "deletionAtTrackedEdge";
  /** Two tracked deletions sharing content: nested, equal or crossing. */
  OVERLAPPING_DELETIONS: "overlappingDeletions";
}> = {
  CUT_AT_TRACKED_CHANGE: "cutAtTrackedChange",
  DELETION_AT_TRACKED_EDGE: "deletionAtTrackedEdge",
  OVERLAPPING_DELETIONS: "overlappingDeletions",
} as const;
export type TrackedOrderPair =
  (typeof TRACKED_ORDER_PAIRS)[keyof typeof TRACKED_ORDER_PAIRS];

/** The edges of a tracked deletion or insertion: where its wrapper starts and ends. */
const trackedEdges = (op: DocumentOp): readonly TextPosition[] => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
      return op.revision === undefined ? [] : [op.from, op.to];
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
      return op.revision === undefined ? [] : [op.at];
    default:
      return [];
  }
};

/**
 * A tracked change's edges in `document`: a tracked edit that meets a record
 * of the same kind, author and date extends that change, so its edges are
 * the merged change's. Without a document, the operation's own.
 */
const changeEdges = (
  op: DocumentOp,
  document: Document | undefined,
): readonly TextPosition[] => {
  const edges = trackedEdges(op);
  const start = edges.at(0);
  const end = edges.at(-1);
  if (
    document === undefined ||
    start === undefined ||
    end === undefined ||
    !("revision" in op) ||
    op.revision === undefined
  )
    return edges;
  const body = findStoryBody(document, start.story);
  const paragraph =
    body === undefined
      ? undefined
      : findStoryParagraph(body, start.blockId)?.paragraph;
  if (paragraph === undefined) return edges;
  const kind =
    op.type === DOCUMENT_OP_TYPES.DELETE_RANGE ? "deletion" : "insertion";
  const { author, date } = op.revision;
  const stamped = leafSpans(paragraph.content)
    .filter(({ ancestors }) =>
      ancestors.some(
        (node) =>
          isTrackedWrapper(node) &&
          node.type === kind &&
          node.info.author === author &&
          node.info.date === date,
      ),
    )
    .map(({ before, after }) => [before.offset, after.offset] as const);
  let from = start.offset;
  let to = end.offset;
  for (let grown = true; grown;) {
    grown = false;
    for (const [left, right] of stamped)
      if (left <= to && right >= from && (left < from || right > to)) {
        from = Math.min(from, left);
        to = Math.max(to, right);
        grown = true;
      }
  }
  return [
    { ...start, offset: from },
    { ...end, offset: to },
  ];
};

/**
 * Where an operation may cut or merge inline records. A direct deletion
 * keeps the two ends of a record it passes through as one, and cuts nothing,
 * unless it joins records as an insertion's inverse.
 */
const cutPoints = (op: DocumentOp): readonly TextPosition[] => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
      if (op.revision !== undefined) return [op.from, op.to];
      return (op.join ?? 0) > 0 ? [op.from] : [];
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return [op.from, op.to];
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
      return [op.at];
    default:
      return [];
  }
};

/**
 * Whether a cut lies at a tracked change's edge or inside its range: the
 * depth of the records running across the whole range, which removing the
 * change joins again, depends on it as much as an edge's does.
 */
const reaches = (
  edges: readonly TextPosition[],
  cuts: readonly TextPosition[],
): boolean => {
  const start = edges.at(0);
  const end = edges.at(-1);
  if (start === undefined || end === undefined) return false;
  return cuts.some(
    (cut) =>
      sameParagraph(start, cut) &&
      start.offset <= cut.offset &&
      cut.offset <= end.offset,
  );
};

// Whether a direct deletion reaches a tracked edge: it removes what lay
// against the edge in one order and moves the edge in the other, and
// covering the change whole leaves the records around it cut in one order.
const reachesEdge = (
  edges: readonly TextPosition[],
  deletion: DocumentOp,
): boolean => {
  if (
    deletion.type !== DOCUMENT_OP_TYPES.DELETE_RANGE ||
    deletion.revision !== undefined
  )
    return false;
  const reached = edges.filter(
    (edge) =>
      sameParagraph(edge, deletion.from) &&
      deletion.from.offset <= edge.offset &&
      edge.offset <= deletion.to.offset,
  ).length;
  return reached > 0;
};

const overlappingDeletions = (left: DocumentOp, right: DocumentOp): boolean =>
  left.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
  right.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
  left.revision !== undefined &&
  right.revision !== undefined &&
  sameParagraph(left.from, right.from) &&
  left.from.offset < right.to.offset &&
  right.from.offset < left.to.offset;

/**
 * The kind of order-dependent tracked pair two concurrent operations form,
 * or undefined when they are outside this policy. Scoped source restorations
 * keep their exclusive edit contract. Symmetric: either may be the tracked one.
 * Given the document both address, a tracked edit's edges are those of the
 * change it extends there (see {@link changeEdges}).
 */
export const trackedOrderPair = (
  op: DocumentOp,
  over: DocumentOp,
  document?: Document,
): TrackedOrderPair | undefined => {
  if (
    (op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT && !isInlineInsertion(op)) ||
    (over.type === DOCUMENT_OP_TYPES.INSERT_CONTENT && !isInlineInsertion(over))
  )
    return undefined;
  if (fragmentedDeletionOverlap(op, over)) return undefined;
  if (overlappingDeletions(op, over))
    return TRACKED_ORDER_PAIRS.OVERLAPPING_DELETIONS;
  const opEdges = changeEdges(op, document);
  const overEdges = changeEdges(over, document);
  if (reachesEdge(opEdges, over) || reachesEdge(overEdges, op))
    return TRACKED_ORDER_PAIRS.DELETION_AT_TRACKED_EDGE;
  if (reaches(opEdges, cutPoints(over)) || reaches(overEdges, cutPoints(op)))
    return TRACKED_ORDER_PAIRS.CUT_AT_TRACKED_CHANGE;
  return undefined;
};
