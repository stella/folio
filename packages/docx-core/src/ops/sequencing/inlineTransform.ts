/**
 * Sequencing of inline insertions and of inline record splits and joins.
 *
 * An inline insertion (a closed slice of units, {@link isInlineInsertion})
 * moves the offsets after its point by its width. At its own offset the
 * zero-width gaps decide; at the same gap the later-sequenced side goes after
 * the earlier one. A range that starts at the insertion point moves past the
 * inserted units and a range that ends there stays before them: the inserted
 * runs carry their own formatting, so neither a concurrent format nor a
 * concurrent deletion takes them in. A range the insertion falls strictly
 * inside is split around it.
 *
 * Inline splits and joins only move record boundaries, so offsets map one to
 * one across them. A concurrent edit that cuts or merges records at the same
 * offset would leave other boundaries in the two orders, and refuses.
 *
 * A seam join (`ifAlike`, as typing and other edits join their runs to alike
 * neighbours) merges only alike records that carry no ids, and changes
 * nothing elsewhere. It moves like a point: at a concurrent insertion's
 * point the sequence order decides, so the seams of typing at one point
 * follow their own runs. An insertion, a format, a deletion or another seam
 * join at its offset passes it: whichever order applies, alike runs at every
 * edited seam end up joined and unlike ones apart.
 */
import { Result } from "better-result";

import {
  type DeleteRangeOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertContentOp,
  type SetRunPropsOp,
  type TextPosition,
} from "../types";
import {
  BATCH_REJECTION_REASONS,
  BatchRejection,
  type SequencedOpEffect,
} from "./envelope";
import {
  compareGapsAtOffset,
  insertionWidth,
  isInlineInsertion,
  mapPoint,
  type PointMapping,
  sameParagraph,
} from "./points";

type Pair = {
  op: DocumentOp;
  over: DocumentOp;
  effect: SequencedOpEffect | undefined;
  order: "before" | "after";
};

type Transformed = Result<readonly DocumentOp[], BatchRejection>;

const refusal = (op: DocumentOp, over: DocumentOp, message: string) =>
  Result.err(
    new BatchRejection({
      reason: BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
      message,
      opType: op.type,
      overType: over.type,
    }),
  );

const isInlineCut = (
  op: DocumentOp,
): op is Extract<DocumentOp, { type: "splitInline" | "joinInline" }> =>
  op.type === DOCUMENT_OP_TYPES.SPLIT_INLINE ||
  op.type === DOCUMENT_OP_TYPES.JOIN_INLINE;

/** A join that merges only alike records at an edit's seam. */
const isSeamJoin = (op: DocumentOp): boolean =>
  op.type === DOCUMENT_OP_TYPES.JOIN_INLINE && op.ifAlike === true;

/**
 * Two positions of one paragraph in document order: by offset, then by the
 * zero-width gap. Undefined when the gaps cannot be compared.
 */
const comparePositions = (
  left: TextPosition,
  right: TextPosition,
): -1 | 0 | 1 | undefined => {
  if (left.offset !== right.offset) return left.offset < right.offset ? -1 : 1;
  return compareGapsAtOffset(left, right);
};

/** A point past an insertion at or before it. */
const past = (point: TextPosition, over: InsertContentOp): TextPosition =>
  point.offset === over.at.offset
    ? {
        ...point,
        offset: point.offset + insertionWidth(over),
        ...(point.zeroWidthBefore === undefined
          ? {}
          : {
              zeroWidthBefore:
                point.zeroWidthBefore - (over.at.zeroWidthBefore ?? 0),
            }),
      }
    : { ...point, offset: point.offset + insertionWidth(over) };

/** Whether a range's ends name zero-width leaves. */
const namesZeroWidth = (op: DeleteRangeOp): boolean =>
  (op.from.zeroWidthBefore ?? 0) !== 0 || (op.to.zeroWidthBefore ?? 0) !== 0;

/** Ranges whose preconditions or stamps need the range unchanged. */
const structural = (op: DeleteRangeOp | SetRunPropsOp): boolean =>
  op.type === DOCUMENT_OP_TYPES.DELETE_RANGE
    ? op.expected !== undefined || op.join !== undefined
    : op.expected !== undefined ||
      op.joinStart !== undefined ||
      op.joinEnd !== undefined;

/** A format split around content inserted strictly inside it needs one record per run. */
const splittable = (op: DeleteRangeOp | SetRunPropsOp): boolean =>
  op.type === DOCUMENT_OP_TYPES.DELETE_RANGE
    ? op.revision === undefined
    : op.revision === undefined && op.propertyReview === undefined;

/** Where a point lies against a range: before it, at an end, inside it or after it. */
const within = (
  point: TextPosition,
  range: DeleteRangeOp | SetRunPropsOp,
): { start: -1 | 0 | 1; end: -1 | 0 | 1 } | undefined => {
  const start = comparePositions(point, range.from);
  const end = comparePositions(point, range.to);
  return start === undefined || end === undefined ? undefined : { start, end };
};

/**
 * Move a range of one paragraph across an inline insertion in it. The range
 * keeps out of the inserted content at both ends, and is split around
 * content inserted strictly inside it.
 */
const rangeOverInsertion = (
  op: DeleteRangeOp | SetRunPropsOp,
  over: InsertContentOp,
): Transformed => {
  const at = over.at;
  const place = within(at, op);
  if (place === undefined)
    return refusal(op, over, "The insertion's zero-width gap is ambiguous.");
  const inside = place.start > 0 && place.end < 0;
  const touches = place.start >= 0 && place.end <= 0;
  if (structural(op) && touches)
    return refusal(
      op,
      over,
      "Structural inverse preconditions require an unchanged range.",
    );
  if (
    op.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
    (namesZeroWidth(op) ||
      ((at.zeroWidthBefore ?? 0) !== 0 && touches && !op.revision))
  )
    return refusal(
      op,
      over,
      "Tracked deletion and zero-width boundaries require an unchanged range.",
    );
  // An empty range is a point: it moves past content inserted before it.
  if (comparePositions(op.from, op.to) === 0)
    return Result.ok([
      place.start < 0
        ? { ...op, from: past(op.from, over), to: past(op.to, over) }
        : op,
    ]);
  if (place.end >= 0) return Result.ok([op]);
  if (!inside)
    return Result.ok([
      { ...op, from: past(op.from, over), to: past(op.to, over) },
    ]);
  if (
    !splittable(op) ||
    at.zeroWidthBefore === undefined ||
    (op.type === DOCUMENT_OP_TYPES.DELETE_RANGE && over.revision)
  )
    return refusal(
      op,
      over,
      op.type === DOCUMENT_OP_TYPES.DELETE_RANGE && op.revision
        ? "Tracking a concurrent insertion requires a revision partition."
        : "Splitting this range around an insertion requires new identities.",
    );
  const first = {
    ...op,
    to: { ...op.to, offset: at.offset, zeroWidthBefore: at.zeroWidthBefore },
  };
  const second = {
    ...op,
    from: {
      ...op.from,
      offset: at.offset + insertionWidth(over),
      zeroWidthBefore: 0,
    },
    to: past(op.to, over),
  };
  // A deletion removes the later part first, so the earlier keeps its offsets.
  return Result.ok(
    op.type === DOCUMENT_OP_TYPES.DELETE_RANGE
      ? [second, first]
      : [first, second],
  );
};

/**
 * Whether a sequenced insertion's effect says it cut an identified record.
 * A batch's own insertion, which a foreign operation is moved across to
 * address the batch's next operation, has no effect and decides no ids.
 */
const cutIdentified = (effect: SequencedOpEffect | undefined): boolean =>
  effect?.type === DOCUMENT_OP_TYPES.INSERT_CONTENT && effect.cutIdentified;

/** A point operation (insertion, split) moved across an inline insertion. */
const pointOverInsertion = (
  op: Extract<
    DocumentOp,
    {
      type:
        | "insertText"
        | "insertContent"
        | "splitBlock"
        | "splitInline"
        | "joinInline";
    }
  >,
  { over, effect, order }: Pair & { over: InsertContentOp },
): Transformed => {
  const tie =
    sameParagraph(op.at, over.at) &&
    op.at.offset === over.at.offset &&
    compareGapsAtOffset(op.at, over.at) === 0;
  const seam = isSeamJoin(op);
  if (
    sameParagraph(op.at, over.at) &&
    op.at.offset === over.at.offset &&
    op.type !== DOCUMENT_OP_TYPES.INSERT_TEXT &&
    op.type !== DOCUMENT_OP_TYPES.INSERT_CONTENT &&
    !seam
  )
    return refusal(
      op,
      over,
      "Cutting records at a concurrent insertion's point needs explicit affinity.",
    );
  if (
    op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK &&
    sameParagraph(op.at, over.at) &&
    ((op.at.zeroWidthBefore ?? 0) !== 0 || (over.at.zeroWidthBefore ?? 0) !== 0)
  )
    return refusal(
      op,
      over,
      "Splitting beside zero-width leaves requires an unchanged paragraph.",
    );
  // Either side may have cut an identified record at the shared gap, and the
  // record's second half would take the id of whichever applied first. A
  // seam join leaves identified records apart.
  if (tie && !seam && cutIdentified(effect))
    return refusal(
      op,
      over,
      "Insertions at one gap that cut an identified record need explicit affinity.",
    );
  const mapping: PointMapping = {
    over,
    effect,
    order,
    splitTiesToSecond: true,
  };
  const at = mapPoint(op.at, "after", mapping);
  return at === undefined
    ? refusal(op, over, "The insertion's zero-width gap is ambiguous.")
    : Result.ok([{ ...op, at }]);
};

/** Refusals that keep the pair symmetric: what refuses one way refuses the other. */
const insertionOverRange = (
  op: InsertContentOp,
  over: DeleteRangeOp | SetRunPropsOp,
): BatchRejection | undefined => {
  if (!sameParagraph(op.at, over.from)) return undefined;
  const place = within(op.at, over);
  if (place === undefined)
    return new BatchRejection({
      reason: BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
      message: "The insertion's zero-width gap is ambiguous.",
      opType: op.type,
      overType: over.type,
    });
  const inside = place.start > 0 && place.end < 0;
  const touches = place.start >= 0 && place.end <= 0;
  const refuse = (message: string) =>
    new BatchRejection({
      reason: BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
      message,
      opType: op.type,
      overType: over.type,
    });
  if (structural(over) && touches)
    return refuse(
      "Structural inverse preconditions require an unchanged range.",
    );
  if (over.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS)
    return inside && !splittable(over)
      ? refuse(
          "Splitting this range around an insertion requires new identities.",
        )
      : undefined;
  if (namesZeroWidth(over))
    return refuse(
      "Tracked deletion and zero-width boundaries require an unchanged range.",
    );
  if (over.revision)
    return inside
      ? refuse(
          "Insertion inside tracked-deleted content requires wrapper affinity.",
        )
      : undefined;
  if (inside && (op.revision || op.at.zeroWidthBefore === undefined))
    return refuse("Splitting this deletion requires new revision identities.");
  if (touches && (op.at.zeroWidthBefore ?? 0) !== 0)
    return refuse(
      "Tracked deletion and zero-width boundaries require an unchanged range.",
    );
  return undefined;
};

/** Whether a range reaches the offset where inline records are cut or merged. */
const rangeReaches = (
  range: DeleteRangeOp | SetRunPropsOp,
  point: TextPosition,
): boolean => {
  if (!sameParagraph(range.from, point)) return false;
  // A deletion removes the records around the point; a format cuts runs at
  // its ends.
  return range.type === DOCUMENT_OP_TYPES.DELETE_RANGE
    ? range.from.offset <= point.offset && point.offset <= range.to.offset
    : range.from.offset === point.offset || range.to.offset === point.offset;
};

const sameOffset = (left: TextPosition, right: TextPosition): boolean =>
  sameParagraph(left, right) && left.offset === right.offset;

// Whether a point is the start of the paragraph a join appends.
const atJoinedStart = (
  point: TextPosition,
  join: Extract<DocumentOp, { type: "joinBlocks" }>,
): boolean =>
  point.offset === 0 &&
  sameParagraph(point, { story: join.story, blockId: join.nextBlockId });

/**
 * Decide a pair in which either side is an inline insertion or an inline
 * split or join. Undefined leaves the pair to the rest of the position
 * transform.
 */
export const transformInlinePair = (pair: Pair): Transformed | undefined => {
  const { op, over, effect, order } = pair;
  if (isInlineInsertion(over)) {
    switch (op.type) {
      case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
        return Result.ok([op]);
      case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
        // Content inserted at the start of the second paragraph meets the
        // first one's last run at the retained boundary.
        if (
          !op.revision &&
          (op.depth ?? 0) === 0 &&
          sameParagraph(over.at, {
            story: op.story,
            blockId: op.nextBlockId,
          }) &&
          over.at.offset === 0
        )
          return refusal(
            op,
            over,
            "Insertion at a retained run boundary needs explicit affinity.",
          );
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
          return undefined;
        return pointOverInsertion(op, { ...pair, over });
      case DOCUMENT_OP_TYPES.DELETE_RANGE:
      case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
        if (!sameParagraph(op.from, over.at)) return Result.ok([op]);
        return rangeOverInsertion(op, over);
      default:
        return undefined;
    }
  }
  if (isInlineInsertion(op)) {
    switch (over.type) {
      case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
        return Result.ok([op]);
      case DOCUMENT_OP_TYPES.DELETE_RANGE:
      case DOCUMENT_OP_TYPES.SET_RUN_PROPS: {
        const refused = insertionOverRange(op, over);
        if (refused !== undefined) return Result.err(refused);
        break;
      }
      case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
        if (sameOffset(op.at, over.at))
          return refusal(
            op,
            over,
            "Cutting records at a concurrent insertion's point needs explicit affinity.",
          );
        break;
      case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
        if (
          !over.revision &&
          (over.depth ?? 0) === 0 &&
          sameParagraph(op.at, {
            story: over.story,
            blockId: over.nextBlockId,
          }) &&
          op.at.offset === 0
        )
          return refusal(
            op,
            over,
            "Insertion at a retained run boundary needs explicit affinity.",
          );
        break;
      case DOCUMENT_OP_TYPES.SPLIT_INLINE:
      case DOCUMENT_OP_TYPES.JOIN_INLINE:
        // Content inserted at a seam join's point cuts the joined runs, and
        // its own seams join them again.
        if (sameOffset(op.at, over.at) && !isSeamJoin(over))
          return refusal(
            op,
            over,
            "Cutting records at a concurrent insertion's point needs explicit affinity.",
          );
        return Result.ok([op]);
      case DOCUMENT_OP_TYPES.INSERT_TEXT:
        break;
      default:
        return undefined;
    }
    const at = mapPoint(op.at, "after", {
      over,
      effect,
      order,
      splitTiesToSecond: true,
    });
    return at === undefined
      ? refusal(
          op,
          over,
          "Position cannot be mapped without structural context.",
        )
      : Result.ok([{ ...op, at }]);
  }
  if (isInlineCut(op)) {
    const seam = isSeamJoin(op);
    switch (over.type) {
      case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
        return Result.ok([op]);
      case DOCUMENT_OP_TYPES.INSERT_TEXT:
      case DOCUMENT_OP_TYPES.SPLIT_INLINE:
      case DOCUMENT_OP_TYPES.JOIN_INLINE:
        if (sameOffset(op.at, over.at)) {
          // Two seam joins at one point: the second finds the runs joined.
          if (seam && isSeamJoin(over)) return Result.ok([op]);
          return refusal(
            op,
            over,
            "Cutting records at a concurrent edit's point needs explicit affinity.",
          );
        }
        break;
      case DOCUMENT_OP_TYPES.DELETE_RANGE:
        if (namesZeroWidth(over) || (!seam && rangeReaches(over, op.at)))
          return refusal(
            op,
            over,
            "A concurrent deletion reached the records this operation cuts.",
          );
        break;
      case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
        if (!seam && rangeReaches(over, op.at))
          return refusal(
            op,
            over,
            "A concurrent format cut runs where this operation cuts records.",
          );
        return Result.ok([op]);
      case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
        // A seam join elsewhere in a split paragraph moves with its half.
        if (seam && !sameOffset(op.at, over.at)) break;
        return refusal(
          op,
          over,
          "Cutting inline records of a split or joined paragraph requires an exclusive edit.",
        );
      case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
        // At the start of the appended paragraph the join keeps the two
        // paragraphs' runs apart.
        if (seam && !atJoinedStart(op.at, over)) break;
        return refusal(
          op,
          over,
          "Cutting inline records of a split or joined paragraph requires an exclusive edit.",
        );
      default:
        return undefined;
    }
    const at = mapPoint(op.at, "after", {
      over,
      effect,
      order,
      splitTiesToSecond: true,
    });
    return at === undefined
      ? refusal(
          op,
          over,
          "Position cannot be mapped without structural context.",
        )
      : Result.ok([{ ...op, at }]);
  }
  if (isInlineCut(over)) {
    const seam = isSeamJoin(over);
    switch (op.type) {
      case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
        return Result.ok([op]);
      case DOCUMENT_OP_TYPES.INSERT_TEXT:
      case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
        return sameOffset(op.at, over.at)
          ? refusal(
              op,
              over,
              "An edit where a concurrent operation cut records needs explicit affinity.",
            )
          : Result.ok([op]);
      case DOCUMENT_OP_TYPES.DELETE_RANGE:
      case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
        return (!seam && rangeReaches(op, over.at)) ||
          (op.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
            namesZeroWidth(op) &&
            sameParagraph(op.from, over.at))
          ? refusal(
              op,
              over,
              "A range reaching records a concurrent operation cut needs captured boundaries.",
            )
          : Result.ok([op]);
      case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
        if (seam && !atJoinedStart(over.at, op)) return Result.ok([op]);
        return refusal(
          op,
          over,
          "Joining a paragraph whose inline records were cut requires an exclusive edit.",
        );
      default:
        return undefined;
    }
  }
  return undefined;
};
