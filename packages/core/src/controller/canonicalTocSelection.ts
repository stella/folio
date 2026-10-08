import { panic } from "better-result";
import {
  DOCUMENT_OP_TYPES,
  SPLIT_HALVES,
  sameStory,
  idKey,
  compareGaps,
  mapTocBookmarkPosition,
  type DocumentOp,
  type TextPosition,
} from "@stll/docx-core/ops";

type TocSelectionMapping = { at: TextPosition; after: TextPosition; ops: readonly DocumentOp[] };

/** TOC insertion retains the selected text and maps its endpoints through the canonical split. */
export const mapTocSelection = (
  point: TextPosition,
  { at, after, ops }: TocSelectionMapping,
): TextPosition => {
  if (!sameStory(point.story, at.story)) return point;
  point = mapTocBookmarkPosition(point, ops);
  at = mapTocBookmarkPosition(at, ops);
  if (idKey(point.blockId) !== idKey(at.blockId)) return point;
  const split = ops.find((op) => op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK);
  if (split === undefined) return point;
  if (point.offset === at.offset && (point.zeroWidthBefore ?? 0) === (at.zeroWidthBefore ?? 0))
    return after;
  if (split.type !== DOCUMENT_OP_TYPES.SPLIT_BLOCK || split.newHalf !== SPLIT_HALVES.FIRST)
    return panic("TOC insertion must retain its following paragraph identity.");
  const pointGap = { offset: point.offset, zeroWidthBefore: point.zeroWidthBefore ?? 0 };
  const splitGap = { offset: split.at.offset, zeroWidthBefore: split.at.zeroWidthBefore ?? 0 };
  if (compareGaps(pointGap, splitGap) < 0) return { ...point, blockId: split.newBlockId };
  if (point.offset === split.at.offset)
    return {
      ...point,
      offset: 0,
      zeroWidthBefore: pointGap.zeroWidthBefore - splitGap.zeroWidthBefore,
    };
  return { ...point, offset: point.offset - split.at.offset };
};
