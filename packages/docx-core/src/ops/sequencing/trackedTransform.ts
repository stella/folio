/**
 * Concurrent tracked changes and the record cuts beside them.
 *
 * A tracked deletion or insertion records, on its wrapper, how many records
 * it cut at each edge, so that rejecting it joins them again. Whether a
 * record runs across an edge depends on what was applied before: another
 * edit that cuts records at the same offset (a tracked edge, an insertion
 * point, a format's end, an inline split or join, a paragraph split) leaves
 * the cut to whichever applies first, and the two orders record different
 * joins. So does a cut inside the range, which decides whether a record
 * runs across the whole of it. A direct deletion reaching one edge of a
 * tracked change and not the other removes what lay against that edge in
 * one order and moves the edge to where it ends in the other. Two tracked deletions that overlap leave
 * the shared content to whichever wraps it first. Such pairs are refused;
 * tracked changes away from every concurrent cut commute.
 */
import { Result } from "better-result";

import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type TextPosition,
} from "../types";
import { BATCH_REJECTION_REASONS, BatchRejection } from "./envelope";
import { sameParagraph } from "./points";

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

// The rule in one direction: the tracked side is `op`.
const conflictsOneWay = (op: DocumentOp, over: DocumentOp): boolean => {
  const edges = trackedEdges(op);
  return reaches(edges, cutPoints(over)) || crossesEdges(edges, over);
};

/**
 * Refuse a pair in which a concurrent cut lies at a tracked change's edge or
 * inside its range, a direct deletion reaches one edge of a tracked change
 * and not the other, or
 * two tracked deletions overlap. Undefined leaves the pair to the position
 * transform. The rule is symmetric, so both orders refuse alike.
 */
export const trackedCutConflict = (
  op: DocumentOp,
  over: DocumentOp,
): Result<never, BatchRejection> | undefined => {
  if (
    !conflictsOneWay(op, over) &&
    !conflictsOneWay(over, op) &&
    !overlappingDeletions(op, over)
  )
    return undefined;
  return Result.err(
    new BatchRejection({
      reason: BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
      message:
        "A tracked change meeting a concurrent cut or tracked deletion needs captured boundaries.",
      opType: op.type,
      overType: over.type,
    }),
  );
};
