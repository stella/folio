import type { Node as PMNode } from "prosemirror-model";

import { OP_STORIES, type TextPosition } from "@stll/docx-core/ops";

import { COMMENT_REFERENCE_NODE_NAME } from "../prosemirror/extensions/nodes/CommentReferenceExtension";
import { TEXT_BOX_ANCHOR_NODE_NAME } from "../prosemirror/extensions/nodes/TextBoxAnchorExtension";
import { findParagraphByParaId } from "../prosemirror/utils/findParagraphByParaId";
import { isZeroWidthAnchor } from "../prosemirror/zeroWidthAnchors";

/**
 * How many logical units an inline leaf takes, as the engine counts them:
 * text its characters, a zero-width anchor (a bookmark or range boundary)
 * none, a comment reference, a text-box anchor and any other atom one.
 */
const leafUnits = (node: PMNode): number => {
  if (node.isText) return node.text?.length ?? 0;
  if (
    isZeroWidthAnchor(node) &&
    node.type.name !== COMMENT_REFERENCE_NODE_NAME &&
    node.type.name !== TEXT_BOX_ANCHOR_NODE_NAME
  )
    return 0;
  return 1;
};

/**
 * Where a body position shows in an editor document that is not a canonical
 * projection: the paragraph whose `paraId` is the position's block id, at the
 * position's offset counted in the engine's logical units (see `leafUnits`),
 * so zero-width anchors before the point do not move it. Null for another
 * story, an unknown paragraph, or an offset past the paragraph's end.
 */
export const legacyEditorPositionAt = (
  doc: PMNode,
  position: TextPosition,
): number | null => {
  if (position.story !== OP_STORIES.MAIN) return null;
  if (!Number.isInteger(position.offset) || position.offset < 0) return null;
  const paragraph = findParagraphByParaId(doc, position.blockId);
  if (paragraph === null) return null;
  const contentStart = paragraph.from + 1;
  let remaining = position.offset;
  let found: number | null = remaining === 0 ? contentStart : null;
  paragraph.node.descendants((node, offset) => {
    if (found !== null) return false;
    if (!node.isLeaf) return true;
    const units = leafUnits(node);
    // A zero-width anchor at the offset is passed: the point lies after it.
    if (units > 0 && remaining <= units) {
      found = contentStart + offset + (node.isText ? remaining : units);
      return false;
    }
    remaining -= units;
    return false;
  });
  return found;
};
