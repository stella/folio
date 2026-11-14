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
 * it. A direct deletion reaching one edge of a tracked change and not the
 * other removes what lay against that edge in one order and moves the edge
 * in the other. Two tracked deletions that overlap leave the shared content
 * to whichever wraps it first: the one sequenced first.
 *
 * Such pairs are accepted and rebased like any other: the authority's order
 * decides the document, and every peer reaches it by replaying its pending
 * batches over the sequenced ones. Applied in the other order they record
 * other cuts, so a peer holding one of them pending gets the document as a
 * replacement rather than a rebased change. These are the only pairs
 * accepted without commuting.
 */
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type TextPosition,
} from "../types";
import { sameParagraph } from "./points";

/** The kinds of concurrent pair whose two orders record different cuts. */
export const TRACKED_ORDER_PAIRS = {
  /** A cut at a tracked change's edge or inside its range. */
  CUT_AT_TRACKED_CHANGE: "cutAtTrackedChange",
  /** A direct deletion reaching one edge of a tracked change and not the other. */
  DELETION_AT_ONE_EDGE: "deletionAtOneEdge",
  /** Two tracked deletions sharing content: nested, equal or crossing. */
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

// Whether a direct deletion reaches some of the tracked edges and not all.
const crossesEdges = (
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
  return reached > 0 && reached < edges.length;
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
 * or undefined when they commute. Symmetric: either may be the tracked one.
 */
export const trackedOrderPair = (
  op: DocumentOp,
  over: DocumentOp,
): TrackedOrderPair | undefined => {
  if (overlappingDeletions(op, over))
    return TRACKED_ORDER_PAIRS.OVERLAPPING_DELETIONS;
  const opEdges = trackedEdges(op);
  const overEdges = trackedEdges(over);
  if (crossesEdges(opEdges, over) || crossesEdges(overEdges, op))
    return TRACKED_ORDER_PAIRS.DELETION_AT_ONE_EDGE;
  if (reaches(opEdges, cutPoints(over)) || reaches(overEdges, cutPoints(op)))
    return TRACKED_ORDER_PAIRS.CUT_AT_TRACKED_CHANGE;
  return undefined;
};
