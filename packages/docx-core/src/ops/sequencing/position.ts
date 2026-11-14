import { DOCUMENT_OP_TYPES, OP_STORIES, type TextPosition } from "../types";
import { restoreDocumentOp } from "../wire";
import { sameStory } from "./address";
import {
  isAnchoredComment,
  positionPastCommentAnchor,
} from "./commentTransform";
import type { SequencedBatch } from "./envelope";
import { isInlineInsertion, mapPoint } from "./points";
import { operationStory, operationTargets } from "./transform";

/**
 * Where a position lands after an ordered journal tail, by paragraph identity
 * and offset, never by an editor's own coordinates. At a foreign insertion's
 * or split's own offset, `before` keeps the position ahead of what was
 * inserted and `after` moves it past. Undefined when the tail removes the
 * position's paragraph or changes it in a way an offset cannot follow.
 */
export const transformPosition = (
  position: TextPosition,
  over: readonly SequencedBatch[],
  affinity: "before" | "after" = "after",
): TextPosition | undefined => {
  let point = position;
  for (const sequenced of over) {
    for (const [index, wire] of sequenced.ops.entries()) {
      const restored = restoreDocumentOp(wire);
      if (restored.isErr()) return undefined;
      const op = restored.value;
      const effect = sequenced.effects?.at(index);
      const story = operationStory(op);
      if (story !== undefined && !sameStory(story, point.story)) continue;
      if (isAnchoredComment(op)) {
        point = positionPastCommentAnchor(point, op, affinity);
        continue;
      }
      if (effect?.type === "touchedBlocks") {
        if (effect.blockIds.includes(point.blockId)) return undefined;
        continue;
      }
      const addressed = operationTargets(op);
      if (addressed !== undefined && !addressed.includes(point.blockId))
        continue;
      switch (op.type) {
        case DOCUMENT_OP_TYPES.INSERT_TEXT:
        case DOCUMENT_OP_TYPES.INSERT_CONTENT:
        case DOCUMENT_OP_TYPES.DELETE_RANGE:
        case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
        case DOCUMENT_OP_TYPES.JOIN_BLOCKS: {
          // Content merged into the records at its point leaves no offset
          // to follow in its paragraph.
          if (
            op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT &&
            !isInlineInsertion(op)
          )
            return undefined;
          const mapped = mapPoint(point, affinity, {
            over: op,
            effect,
            order: affinity,
            splitTiesToSecond: true,
          });
          if (mapped === undefined) return undefined;
          point = mapped;
          continue;
        }
        // Formatting, inline record splits and joins, paragraph insertion
        // beside the position, and package metadata leave every offset where
        // it was.
        case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
        case DOCUMENT_OP_TYPES.SPLIT_INLINE:
        case DOCUMENT_OP_TYPES.JOIN_INLINE:
        case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
        case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
        case DOCUMENT_OP_TYPES.SET_REVISION_INFO:
        case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
        case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
        case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
        case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
        case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
        case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
        case DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES:
          continue;
        // A header or footer part changes no body offset.
        case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
        case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
        case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
          if (point.story === OP_STORIES.MAIN) continue;
          return undefined;
        case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
        case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
        case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
        case DOCUMENT_OP_TYPES.REPLACE_INLINE:
        case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
        case DOCUMENT_OP_TYPES.INSERT_TABLE:
        case DOCUMENT_OP_TYPES.DELETE_TABLE:
        case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
        case DOCUMENT_OP_TYPES.INSERT_ROW:
        case DOCUMENT_OP_TYPES.DELETE_ROW:
        case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
        case DOCUMENT_OP_TYPES.INSERT_COLUMN:
        case DOCUMENT_OP_TYPES.DELETE_COLUMN:
        case DOCUMENT_OP_TYPES.MERGE_CELLS:
        case DOCUMENT_OP_TYPES.SPLIT_CELL:
        case DOCUMENT_OP_TYPES.SET_TABLE_GRID:
        case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
        case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
        case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
        case DOCUMENT_OP_TYPES.SET_TABLE:
        case DOCUMENT_OP_TYPES.ADD_NOTE:
        case DOCUMENT_OP_TYPES.REMOVE_NOTE:
        // A reply, a deletion or a restoration adds or removes comment
        // references the operation does not name.
        case DOCUMENT_OP_TYPES.CREATE_COMMENT:
        case DOCUMENT_OP_TYPES.DELETE_COMMENT:
        case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
          return undefined;
        default: {
          const exhaustive: never = op;
          return exhaustive;
        }
      }
    }
  }
  return point;
};
