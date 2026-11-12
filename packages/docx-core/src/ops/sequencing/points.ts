/** Moving one text position across one foreign operation. */
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
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
    return {
      ...point,
      offset:
        point.offset <= over.from.offset
          ? point.offset
          : Math.max(
              over.from.offset,
              point.offset - (over.to.offset - over.from.offset),
            ),
    };
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
  if (over.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS) return point;
  return undefined;
};
