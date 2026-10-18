import type { Node as PMNode } from "prosemirror-model";

import { COMMENT_REFERENCE_NODE_NAME } from "../prosemirror/extensions/nodes/CommentReferenceExtension";
import { RANGE_ANCHOR_NODE_NAME } from "../prosemirror/extensions/nodes/RangeAnchorExtension";

/** Whether an inline node locates a comment without being content covered by it. */
export const isCommentAnchorCarrier = (node: PMNode): boolean =>
  node.type.name === COMMENT_REFERENCE_NODE_NAME || node.type.name === RANGE_ANCHOR_NODE_NAME;
