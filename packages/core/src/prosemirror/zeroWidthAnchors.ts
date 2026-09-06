/**
 * Inline nodes that occupy a position and show nothing.
 *
 * A bookmark boundary, a text-box anchor and Word's cached pagination
 * boundary all exist so a position survives an edit. None of them is content:
 * a reader sees no character where one sits, none carries a revision of its
 * own, and the format has no way to say one was inserted or deleted outside a
 * hyperlink.
 *
 * Two questions turn on that, and they used to be answered separately, which
 * is how a deleted paragraph came back as a blank line holding an invisible
 * anchor:
 *
 * - *Does deleting this block have to mark it?* No — marking it produces a
 *   document the serializer refuses.
 * - *Does it keep an emptied paragraph alive?* No — a paragraph left holding
 *   only these is a paragraph whose content is gone.
 */

import type { Node as PMNode } from "prosemirror-model";

import { BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BookmarkBoundaryExtension";
import { RENDERED_PAGE_BREAK_NODE_NAME } from "./extensions/nodes/RenderedPageBreakExtension";
import { TEXT_BOX_ANCHOR_NODE_NAME } from "./extensions/nodes/TextBoxAnchorExtension";

const ZERO_WIDTH_ANCHOR_NODE_NAMES: ReadonlySet<string> = new Set([
  BOOKMARK_BOUNDARY_NODE_NAME,
  RENDERED_PAGE_BREAK_NODE_NAME,
  TEXT_BOX_ANCHOR_NODE_NAME,
]);

export const isZeroWidthAnchor = (node: PMNode): boolean =>
  ZERO_WIDTH_ANCHOR_NODE_NAMES.has(node.type.name);

/**
 * Whether the textblock holds nothing but zero-width anchors, which is what
 * "everything in it was resolved away" looks like in the document.
 */
export const holdsNoContent = (block: PMNode): boolean => {
  let empty = true;
  block.forEach((child) => {
    if (!isZeroWidthAnchor(child)) {
      empty = false;
    }
  });
  return empty;
};
