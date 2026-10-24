/**
 * Inline nodes that occupy a position and show nothing.
 *
 * Bookmark and move-range boundaries, range and text-box anchors, comment
 * references, and Word's cached pagination boundaries all exist so a position
 * survives an edit. None of them is content:
 * a reader sees no character where one sits, none carries a revision of its
 * own, and the format has no way to say one was inserted or deleted outside a
 * hyperlink.
 *
 * - *Does deleting this block have to mark it?* No — marking it produces a
 *   document the serializer refuses.
 * - *Does it keep a paragraph whose mark goes?* That is the engine's one
 *   paragraph-content rule (`paragraphContentVerdict` in
 *   `@stll/docx-core/ops`): a range boundary or comment reference does, so its
 *   bookmark, comment or move survives the mark; a cached page break does not.
 *   Each anchor names the engine item it stands for, so the editor and the
 *   engine decide every paragraph-mark outcome alike.
 */

import type { Node as PMNode } from "prosemirror-model";

import {
  type ParagraphContentVerdict,
  paragraphContentVerdict,
  type ZeroWidthItem,
} from "@stll/docx-core/ops";

import { expectCommentRangeBoundaryAttrs } from "./commentRangeBoundaryAttrs";
import { BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BookmarkBoundaryExtension";
import { COMMENT_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/CommentRangeBoundaryExtension";
import { COMMENT_REFERENCE_NODE_NAME } from "./extensions/nodes/CommentReferenceExtension";
import { MOVE_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/MoveRangeBoundaryExtension";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import { RENDERED_PAGE_BREAK_NODE_NAME } from "./extensions/nodes/RenderedPageBreakExtension";
import { TEXT_BOX_ANCHOR_NODE_NAME } from "./extensions/nodes/TextBoxAnchorExtension";

/**
 * The engine item each anchor stands for: a zero-width item the content rule
 * classifies, or one unit of width (a comment reference is an atom, a text box
 * anchor stands for its drawing in the run stream).
 */
const ANCHOR_ITEMS = new Map<string, ZeroWidthItem | "unit">([
  [BOOKMARK_BOUNDARY_NODE_NAME, "bookmarkStart"],
  [COMMENT_REFERENCE_NODE_NAME, "unit"],
  [COMMENT_RANGE_BOUNDARY_NODE_NAME, "commentRangeStart"],
  [MOVE_RANGE_BOUNDARY_NODE_NAME, "moveFromRangeStart"],
  [RANGE_ANCHOR_NODE_NAME, "commentRangeStart"],
  [RENDERED_PAGE_BREAK_NODE_NAME, "renderedPageBreak"],
  [TEXT_BOX_ANCHOR_NODE_NAME, "unit"],
]);

export const isZeroWidthAnchor = (node: PMNode): boolean =>
  ANCHOR_ITEMS.has(node.type.name);

/**
 * A node's width the way the engine counts it: an atom or leaf (text, a
 * field, an image) is one unit or more, a container (an inline content
 * control) counts its children in place, and an anchor counts as the engine
 * item it stands for, which for a zero-width item is no width.
 */
const collectParagraphContent = (
  node: PMNode,
  zeroWidthItems: ZeroWidthItem[],
): number => {
  if (node.type.name === COMMENT_RANGE_BOUNDARY_NODE_NAME) {
    zeroWidthItems.push(expectCommentRangeBoundaryAttrs(node).marker.type);
    return 0;
  }
  const item = ANCHOR_ITEMS.get(node.type.name);
  if (item !== undefined) {
    if (item === "unit") return 1;
    zeroWidthItems.push(item);
    return 0;
  }
  if (node.isAtom || node.isLeaf) {
    return 1;
  }
  let width = 0;
  node.forEach((child) => {
    width += collectParagraphContent(child, zeroWidthItems);
  });
  return width;
};

/**
 * Whether the textblock holds content its paragraph mark must not take away,
 * by the engine's rule: the paragraph's width, and each anchor directly in
 * it counted as the engine item it stands for.
 */
export const paragraphContentOf = (block: PMNode): ParagraphContentVerdict => {
  let width = 0;
  const zeroWidthItems: ZeroWidthItem[] = [];
  block.forEach((child) => {
    width += collectParagraphContent(child, zeroWidthItems);
  });
  return paragraphContentVerdict({ width, zeroWidthItems });
};
