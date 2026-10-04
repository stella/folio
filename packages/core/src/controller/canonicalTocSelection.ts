import { panic } from "better-result";
import {
  DOCUMENT_OP_TYPES,
  SPLIT_HALVES,
  sameStory,
  idKey,
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
  if (idKey(point.blockId) !== idKey(at.blockId)) return point;
  if (point.offset === at.offset && (point.zeroWidthBefore ?? 0) === (at.zeroWidthBefore ?? 0))
    return after;
  const split = ops.find((op) => op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK);
  if (split === undefined) return point;
  if (split.type !== DOCUMENT_OP_TYPES.SPLIT_BLOCK || split.newHalf !== SPLIT_HALVES.FIRST)
    return panic("TOC insertion must retain its following paragraph identity.");
  if (point.offset < split.at.offset) return { ...point, blockId: split.newBlockId };
  return { ...point, offset: point.offset - split.at.offset };
};
