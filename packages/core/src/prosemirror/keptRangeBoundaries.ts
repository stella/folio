/**
 * Range boundaries rejecting a change keeps from the content it removes.
 *
 * The engine keeps every bookmark and move range boundary of rejected content
 * where that content was (`rangeBoundaries.ts` in `@stll/docx-core/ops`); a
 * comment whose whole range was rejected goes with its comment, and accepting
 * removes what a direct deletion would. The editor resolves the same way:
 * when rejecting, such a boundary loses its revision marks and stays.
 */

import type { Node as PMNode } from "prosemirror-model";

import { BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BookmarkBoundaryExtension";
import { MOVE_RANGE_BOUNDARY_NODE_NAME } from "./extensions/nodes/MoveRangeBoundaryExtension";
import { RANGE_ANCHOR_NODE_NAME } from "./extensions/nodes/RangeAnchorExtension";
import { expectRangeAnchorAttrs } from "./rangeAnchorAttrs";

/** Whether resolving away the content around `node` keeps it in place. */
export const keepsRangeBoundary = (node: PMNode): boolean => {
  switch (node.type.name) {
    case BOOKMARK_BOUNDARY_NODE_NAME:
    case MOVE_RANGE_BOUNDARY_NODE_NAME:
      return true;
    // A point range: a move keeps it, a comment's whole range goes with the comment.
    case RANGE_ANCHOR_NODE_NAME:
      return expectRangeAnchorAttrs(node).start.type !== "commentRangeStart";
    default:
      return false;
  }
};
