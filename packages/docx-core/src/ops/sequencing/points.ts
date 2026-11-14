/** Moving one text position across one foreign operation. */
import { paragraphIdsIn } from "../ids";
import { leafSpans } from "../leaves";
import { contentWidth } from "../offsets";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertContentOp,
  SPLIT_HALVES,
  type TextPosition,
} from "../types";
import { sameStory } from "./address";
import type { SequencedOpEffect } from "./envelope";

type ParagraphAddress = Pick<TextPosition, "story" | "blockId">;
export const sameParagraph = (
  left: ParagraphAddress,
  right: ParagraphAddress,
): boolean =>
  sameStory(left.story, right.story) && left.blockId === right.blockId;

/**
 * Content inserted at one point that only adds units: a closed slice (it
 * merges with nothing at the point) of runs, text and atoms, without
 * zero-width leaves, comment references or paragraphs of its own. It moves
 * the offsets after its point by its width, as text does.
 */
export const isInlineInsertion = (op: DocumentOp): op is InsertContentOp =>
  op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT &&
  op.slice.openStart === 0 &&
  op.slice.openEnd === 0 &&
  contentWidth(op.slice.content) > 0 &&
  paragraphIdsIn(op.slice.content).length === 0 &&
  leafSpans(op.slice.content).every(
    ({ node, before, after }) =>
      after.offset > before.offset && node.type !== "commentReference",
  );

/** The units an inline insertion adds. */
export const insertionWidth = (op: InsertContentOp): number =>
  contentWidth(op.slice.content);

/**
 * How two gaps at one offset compare: by their zero-width leaves when both
 * name them; an absent count is the gap before every zero-width leaf only
 * where the other names none either. Undefined when that cannot be told.
 */
export const compareGapsAtOffset = (
  left: TextPosition,
  right: TextPosition,
): -1 | 0 | 1 | undefined => {
  const mine = left.zeroWidthBefore;
  const theirs = right.zeroWidthBefore;
  if (mine === undefined || theirs === undefined)
    return (mine ?? 0) === 0 && (theirs ?? 0) === 0 ? 0 : undefined;
  if (mine === theirs) return 0;
  return mine < theirs ? -1 : 1;
};

/**
 * A point past an inline insertion at its own offset: after the inserted
 * units, before the zero-width leaves the insertion point had after it.
 */
const pastInsertion = (
  point: TextPosition,
  over: InsertContentOp,
): TextPosition => {
  const moved = { ...point, offset: point.offset + insertionWidth(over) };
  if (point.offset !== over.at.offset || point.zeroWidthBefore === undefined)
    return moved;
  return {
    ...moved,
    zeroWidthBefore: point.zeroWidthBefore - (over.at.zeroWidthBefore ?? 0),
  };
};

/**
 * Move a point across an inline insertion. At the insertion's own offset the
 * zero-width gaps decide; at the same gap the later-sequenced side goes after
 * the earlier one (`order`). Undefined when the gaps cannot be compared.
 */
const mapInsertion = (
  point: TextPosition,
  over: InsertContentOp,
  order: "before" | "after",
): TextPosition | undefined => {
  if (!sameParagraph(point, over.at) || point.offset < over.at.offset)
    return point;
  if (point.offset > over.at.offset) return pastInsertion(point, over);
  const gap = compareGapsAtOffset(point, over.at);
  if (gap === undefined) return undefined;
  if (gap < 0 || (gap === 0 && order === "before")) return point;
  return pastInsertion(point, over);
};

/** The facts a point needs to move across one foreign operation. */
export type PointMapping = {
  over: DocumentOp;
  effect: SequencedOpEffect | undefined;
  order: "before" | "after";
  /** Whether a point at a split's own offset may move into the second half. */
  splitTiesToSecond: boolean;
};

/** Move a point across one foreign operation by paragraph identity and offset. */
export const mapPoint = (
  point: TextPosition,
  affinity: "before" | "after",
  { over, effect, order, splitTiesToSecond }: PointMapping,
): TextPosition | undefined => {
  if (over.type === DOCUMENT_OP_TYPES.INSERT_TEXT) {
    if (!sameParagraph(point, over.at)) return point;
    const shift =
      point.offset > over.at.offset ||
      (point.offset === over.at.offset && order === "after");
    return shift
      ? { ...point, offset: point.offset + over.text.length }
      : point;
  }
  if (over.type === DOCUMENT_OP_TYPES.DELETE_RANGE) {
    if (!sameParagraph(point, over.from)) return point;
    if (over.revision) return point;
    if (
      (point.zeroWidthBefore ?? 0) !== 0 &&
      point.offset >= over.from.offset &&
      point.offset <= over.to.offset
    )
      return undefined;
    if (point.offset <= over.from.offset) return point;
    const offset = Math.max(
      over.from.offset,
      point.offset - (over.to.offset - over.from.offset),
    );
    // A point the deletion closes up to its start lands after the zero-width
    // leaves the range kept there.
    return offset === over.from.offset &&
      point.zeroWidthBefore !== undefined &&
      over.from.zeroWidthBefore !== undefined
      ? { ...point, offset, zeroWidthBefore: over.from.zeroWidthBefore }
      : { ...point, offset };
  }
  if (over.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK) {
    if (!sameParagraph(point, over.at)) return point;
    const half =
      over.newHalf ??
      (effect?.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK
        ? effect.newHalf
        : undefined);
    if (
      !half ||
      (point.zeroWidthBefore ?? 0) !== 0 ||
      (over.at.zeroWidthBefore ?? 0) !== 0
    )
      return undefined;
    const second =
      point.offset > over.at.offset ||
      (point.offset === over.at.offset &&
        affinity === "after" &&
        splitTiesToSecond);
    return {
      ...point,
      blockId:
        second === (half === SPLIT_HALVES.SECOND)
          ? over.newBlockId
          : over.at.blockId,
      offset: second ? point.offset - over.at.offset : point.offset,
    };
  }
  if (over.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS) {
    if (
      !sameStory(point.story, over.story) ||
      (point.blockId !== over.blockId && point.blockId !== over.nextBlockId)
    )
      return point;
    if (over.revision) return point;
    const length =
      effect?.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS
        ? effect.firstLength
        : undefined;
    if (length === undefined) return undefined;
    const survivor =
      over.survivor === SPLIT_HALVES.FIRST ? over.blockId : over.nextBlockId;
    return {
      ...point,
      blockId: survivor,
      offset: point.offset + (point.blockId === over.nextBlockId ? length : 0),
    };
  }
  if (isInlineInsertion(over)) return mapInsertion(point, over, order);
  // Formatting and inline record splits and joins move no offset.
  if (
    over.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS ||
    over.type === DOCUMENT_OP_TYPES.SPLIT_INLINE ||
    over.type === DOCUMENT_OP_TYPES.JOIN_INLINE
  )
    return point;
  return undefined;
};
