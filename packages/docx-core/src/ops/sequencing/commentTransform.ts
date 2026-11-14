/**
 * Sequencing of comment operations.
 *
 * A range or point comment's anchor moves like text over concurrent text and
 * structure operations, and the comment reference it adds (one offset unit at
 * the range's end or at the point) moves the concurrent positions after it.
 * A position at an anchor's own offset has no affinity to follow and refuses.
 * Comment operations of one thread decide deterministically: a pair that
 * commutes passes, any other is refused naming the comments part.
 */
import { Result } from "better-result";

import {
  type CommentAnchor,
  type CommentOp,
  type CreateCommentOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type TextPosition,
} from "../types";
import {
  BATCH_REJECTION_REASONS,
  BatchRejection,
  EXCLUSIVE_PARTS,
  type ExclusivePart,
  type SequencedOpEffect,
} from "./envelope";
import {
  insertionWidth,
  isInlineInsertion,
  mapPoint,
  sameParagraph,
} from "./points";

export const isCommentOp = (op: DocumentOp): op is CommentOp => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
      return true;
    default:
      return false;
  }
};

/** A creation whose anchor names text positions: a range or a point. */
export type AnchoredComment = CreateCommentOp & {
  anchor: Extract<CommentAnchor, { kind: "range" | "point" }>;
};

export const isAnchoredComment = (op: DocumentOp): op is AnchoredComment =>
  op.type === DOCUMENT_OP_TYPES.CREATE_COMMENT &&
  (op.anchor.kind === "range" || op.anchor.kind === "point");

/** A range's two ends, or the point. */
export const anchorPoints = (op: AnchoredComment): TextPosition[] =>
  op.anchor.kind === "range" ? [op.anchor.from, op.anchor.to] : [op.anchor.at];

/** Where the comment reference occupies one offset unit. */
const referenceAt = (op: AnchoredComment): TextPosition =>
  op.anchor.kind === "range" ? op.anchor.to : op.anchor.at;

const withPoints = (
  op: AnchoredComment,
  points: readonly TextPosition[],
): AnchoredComment | undefined => {
  const [first, second] = points;
  if (first === undefined) return undefined;
  if (op.anchor.kind === "point")
    return { ...op, anchor: { kind: "point", at: first } };
  return second === undefined
    ? undefined
    : { ...op, anchor: { kind: "range", from: first, to: second } };
};

type Pair = {
  op: DocumentOp;
  over: DocumentOp;
  effect: SequencedOpEffect | undefined;
  order: "before" | "after";
};

const unsupported = (op: DocumentOp, over: DocumentOp, message: string) =>
  Result.err(
    new BatchRejection({
      reason: BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
      message,
      opType: op.type,
      overType: over.type,
    }),
  );

const partConflict = (
  op: DocumentOp,
  over: DocumentOp,
  part: ExclusivePart,
  message: string,
) =>
  Result.err(
    new BatchRejection({
      reason: BATCH_REJECTION_REASONS.PART_CONFLICT,
      message,
      opType: op.type,
      overType: over.type,
      part,
    }),
  );

const blockConflict = (op: DocumentOp, over: DocumentOp) =>
  Result.err(
    new BatchRejection({
      reason: BATCH_REJECTION_REASONS.BLOCK_CONFLICT,
      message:
        "A concurrent operation changed a paragraph this comment is anchored in.",
      opType: op.type,
      overType: over.type,
    }),
  );

const touchedBlocks = (
  effect: SequencedOpEffect | undefined,
): readonly string[] | undefined =>
  effect?.type === "touchedBlocks" ? effect.blockIds : undefined;

// Whether a point sits at the offset of one of the anchor's own markers.
const atAnchor = (point: TextPosition, anchor: AnchoredComment): boolean =>
  anchorPoints(anchor).some(
    (end) => sameParagraph(point, end) && point.offset === end.offset,
  );

// A point past the comment reference moves past the unit it occupies.
const pastReference = (
  point: TextPosition,
  anchor: AnchoredComment,
): TextPosition => {
  const reference = referenceAt(anchor);
  return sameParagraph(point, reference) && point.offset > reference.offset
    ? { ...point, offset: point.offset + 1 }
    : point;
};

/**
 * Move an anchored comment's anchor across a concurrent text or structure
 * operation. An anchor at the operation's own offset, or inside or at the
 * edge of a concurrent deletion, refuses: a range collapsed by a deletion
 * would become a point comment in one order and stay a range comment in the
 * other.
 */
export const moveCommentAnchor = ({
  op,
  over,
  effect,
  order,
}: Pair & { op: AnchoredComment }): Result<
  readonly DocumentOp[],
  BatchRejection
> => {
  const points = anchorPoints(op);
  const moved = (next: readonly (TextPosition | undefined)[]) => {
    const defined = next.filter(
      (point): point is TextPosition => point !== undefined,
    );
    const rebuilt =
      defined.length === next.length ? withPoints(op, defined) : undefined;
    return rebuilt === undefined
      ? unsupported(op, over, "The comment anchor cannot be mapped.")
      : Result.ok([rebuilt]);
  };
  switch (over.type) {
    // No offset moves: a format splits runs, a paragraph is added beside.
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
      return Result.ok([op]);
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION: {
      const touched = touchedBlocks(effect);
      if (
        touched === undefined ||
        points.some(({ blockId }) => touched.includes(blockId))
      )
        return blockConflict(op, over);
      return Result.ok([op]);
    }
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
      if (
        points.some(
          (point) =>
            sameParagraph(point, over.at) && point.offset === over.at.offset,
        )
      )
        return unsupported(
          op,
          over,
          "Insertion at a comment anchor's own offset needs explicit affinity.",
        );
      return moved(
        points.map((point) =>
          sameParagraph(point, over.at) && point.offset > over.at.offset
            ? { ...point, offset: point.offset + over.text.length }
            : point,
        ),
      );
    case DOCUMENT_OP_TYPES.INSERT_CONTENT: {
      if (!isInlineInsertion(over)) return blockConflict(op, over);
      if (
        points.some(
          (point) =>
            sameParagraph(point, over.at) && point.offset === over.at.offset,
        )
      )
        return unsupported(
          op,
          over,
          "Insertion at a comment anchor's own offset needs explicit affinity.",
        );
      const width = insertionWidth(over);
      return moved(
        points.map((point) =>
          sameParagraph(point, over.at) && point.offset > over.at.offset
            ? { ...point, offset: point.offset + width }
            : point,
        ),
      );
    }
    // Cutting or merging inline records moves no offset; at an anchor's own
    // offset the markers' place among the cut records is not known.
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
      return points.some(
        (point) =>
          sameParagraph(point, over.at) && point.offset === over.at.offset,
      )
        ? unsupported(
            op,
            over,
            "Cutting records at a comment anchor's own offset needs explicit affinity.",
          )
        : Result.ok([op]);
    case DOCUMENT_OP_TYPES.DELETE_RANGE: {
      if (
        points.some(
          (point) =>
            sameParagraph(point, over.from) &&
            over.from.offset <= point.offset &&
            point.offset <= over.to.offset,
        )
      )
        return unsupported(
          op,
          over,
          "A concurrent deletion reached the comment's anchor.",
        );
      // A tracked deletion keeps its text, so no offset moves.
      const length = over.revision ? 0 : over.to.offset - over.from.offset;
      return moved(
        points.map((point) =>
          sameParagraph(point, over.from) && point.offset > over.to.offset
            ? { ...point, offset: point.offset - length }
            : point,
        ),
      );
    }
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS: {
      if (
        over.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK &&
        points.some(
          (point) =>
            sameParagraph(point, over.at) && point.offset === over.at.offset,
        )
      )
        return unsupported(
          op,
          over,
          "A split at a comment anchor's own offset needs explicit affinity.",
        );
      const mapping = { over, effect, order, splitTiesToSecond: true };
      return moved(points.map((point) => mapPoint(point, "after", mapping)));
    }
    default:
      return unsupported(
        op,
        over,
        "This operation pair has no position transform.",
      );
  }
};

/**
 * Move a text or structure operation across a concurrent anchored comment:
 * past its reference when it comes after it, refused at an anchor's own
 * offset and when its range reaches an anchor.
 */
export const passCommentAnchor = ({
  op,
  over,
}: Pair & { over: AnchoredComment }): Result<
  readonly DocumentOp[],
  BatchRejection
> => {
  const reaches = (from: TextPosition, to: TextPosition): boolean =>
    anchorPoints(over).some(
      (point) =>
        sameParagraph(point, from) &&
        from.offset <= point.offset &&
        point.offset <= to.offset,
    );
  switch (op.type) {
    // These address paragraphs, never offsets.
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
      return Result.ok([op]);
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
      if (
        op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT &&
        !isInlineInsertion(op)
      )
        return unsupported(
          op,
          over,
          "This operation pair has no position transform.",
        );
      if (atAnchor(op.at, over))
        return unsupported(
          op,
          over,
          "An edit at a comment anchor's own offset needs explicit affinity.",
        );
      return Result.ok([{ ...op, at: pastReference(op.at, over) }]);
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      if (reaches(op.from, op.to))
        return unsupported(
          op,
          over,
          "A range reaching a concurrent comment's anchor needs captured boundaries.",
        );
      return Result.ok([
        {
          ...op,
          from: pastReference(op.from, over),
          to: pastReference(op.to, over),
        },
      ]);
    default:
      return unsupported(
        op,
        over,
        "This operation pair has no position transform.",
      );
  }
};

/** Where a position lands after an anchored comment, for selections and history. */
export const positionPastCommentAnchor = (
  point: TextPosition,
  anchor: AnchoredComment,
  affinity: "before" | "after",
): TextPosition => {
  let next = point;
  // The markers change the zero-width gaps at the anchor's own offsets.
  if (atAnchor(point, anchor)) {
    const { zeroWidthBefore: _gap, ...offsetOnly } = next;
    next = offsetOnly;
  }
  const reference = referenceAt(anchor);
  if (
    sameParagraph(point, reference) &&
    (point.offset > reference.offset ||
      (point.offset === reference.offset && affinity === "after"))
  )
    next = { ...next, offset: next.offset + 1 };
  return next;
};

/** How an operation uses one comment. */
type Access = "record" | "content" | "status" | "thread";

// The comments an operation creates, edits or removes, and the parent a reply
// joins; a sequenced operation's effect adds the replies a deletion removed.
const accesses = (
  op: CommentOp,
  effect?: SequencedOpEffect,
): Map<number, Access> => {
  const used = new Map<number, Access>();
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
      used.set(op.comment.id, "record");
      if (op.anchor.kind === "reply") used.set(op.anchor.parentId, "thread");
      break;
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
      used.set(op.id, "content");
      break;
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
      used.set(op.id, "status");
      break;
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
      used.set(op.id, "record");
      break;
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
      for (const id of op.ids) used.set(id, "record");
      break;
    default: {
      const exhaustive: never = op;
      return exhaustive;
    }
  }
  if (effect?.type === "touchedBlocks")
    for (const id of effect.comments?.ids ?? [])
      if (!used.has(id)) used.set(id, "record");
  return used;
};

// Whether two uses of one comment commute: the content and the resolution
// are separate fields, a reply only needs its parent to stay, and the same
// resolution twice is one resolution.
const commuting = (
  mine: Access,
  theirs: Access,
  op: CommentOp,
  over: CommentOp,
): boolean => {
  if (mine === "record" || theirs === "record") return false;
  // Replies only need their parent to stay; sibling replies take their
  // places beside the parent's anchors by id.
  if (mine === "thread" || theirs === "thread") return true;
  if (mine !== theirs) return true;
  return (
    mine === "status" &&
    op.type === DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION &&
    over.type === DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION &&
    op.status === over.status
  );
};

// Whether the operation may add a comment part relationship: a reply or a
// resolution needs the extended comment part. Added after a concurrent
// removal it would take another relationship id than it does before it.
const mayAddRelationship = (op: CommentOp): boolean => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
      return op.anchor.kind === "reply" || op.anchor.kind === "revision";
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
      return true;
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
      return false;
    default: {
      const exhaustive: never = op;
      return exhaustive;
    }
  }
};

// A restoration puts records back at the list positions it captured and
// sets the comment part relationships it captured: only a content edit
// leaves both as they were.
const restorationCommutes = (other: CommentOp): boolean =>
  other.type === DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT;

/**
 * Decide a pair of concurrent comment operations. Uses of one comment
 * commute when they edit different fields (content and resolution), when a
 * reply's parent is only edited, resolved or replied to again, and when both
 * set the same resolution; any other use of one comment, a restoration beside anything but a content edit, and a reply or resolution
 * after a concurrent removal of a comment part relationship are refused
 * naming the part. A range or point anchor moves past a concurrent comment's
 * reference.
 */
export const transformCommentPair = ({
  op,
  over,
  effect,
}: Pair & { op: CommentOp; over: CommentOp }): Result<
  readonly DocumentOp[],
  BatchRejection
> => {
  const mine = accesses(op);
  const theirs = accesses(over, effect);
  for (const [id, access] of mine) {
    const other = theirs.get(id);
    if (other !== undefined && !commuting(access, other, op, over))
      return partConflict(
        op,
        over,
        EXCLUSIVE_PARTS.COMMENTS,
        "A concurrent operation changed the comment this operation uses.",
      );
  }
  if (
    effect?.type === "touchedBlocks" &&
    effect.comments?.relationshipRemoved === true &&
    mayAddRelationship(op)
  )
    return partConflict(
      op,
      over,
      EXCLUSIVE_PARTS.RELATIONSHIPS,
      "A concurrent comment operation removed a comment part relationship.",
    );
  if (
    (op.type === DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE &&
      !restorationCommutes(over)) ||
    (over.type === DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE &&
      !restorationCommutes(op))
  )
    return partConflict(
      op,
      over,
      EXCLUSIVE_PARTS.COMMENTS,
      "A comment restoration needs the comment list and parts unchanged.",
    );
  const touched = touchedBlocks(effect) ?? [];
  if (isAnchoredComment(op)) {
    if (isAnchoredComment(over)) {
      const points = anchorPoints(op);
      if (points.some((point) => atAnchor(point, over)))
        return unsupported(
          op,
          over,
          "Two comments anchored at one offset need explicit affinity.",
        );
      const moved = withPoints(
        op,
        points.map((point) => pastReference(point, over)),
      );
      return moved === undefined
        ? unsupported(op, over, "The comment anchor cannot be mapped.")
        : Result.ok([moved]);
    }
    // A reply, deletion or restoration changed anchors found in the document.
    if (anchorPoints(op).some(({ blockId }) => touched.includes(blockId)))
      return blockConflict(op, over);
    return Result.ok([op]);
  }
  if (op.type === DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE) {
    const addressed = new Set(
      [...op.expected.anchors, ...op.state.anchors].map(
        ({ blockId }) => blockId,
      ),
    );
    if (touched.some((blockId) => addressed.has(blockId)))
      return blockConflict(op, over);
  }
  // Replies, edits, resolutions and deletions find their anchors in the
  // document they apply to.
  return Result.ok([op]);
};
