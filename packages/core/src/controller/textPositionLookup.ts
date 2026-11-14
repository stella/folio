import type { Node as PMNode } from "prosemirror-model";

import { OP_STORIES, type TextPosition } from "@stll/docx-core/ops";

import { findParagraphByParaId } from "../prosemirror/utils/findParagraphByParaId";

/**
 * Where a body position shows in an editor document that is not a canonical
 * projection: the paragraph whose `paraId` is the position's block id, at the
 * position's offset counted over the paragraph's text, each inline leaf that
 * is not text counting as one unit. Null for another story, an unknown
 * paragraph, or an offset past the paragraph's end.
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
    const units = node.isText ? (node.text?.length ?? 0) : 1;
    if (remaining <= units) {
      found = contentStart + offset + (node.isText ? remaining : units);
      return false;
    }
    remaining -= units;
    return false;
  });
  return found;
};
