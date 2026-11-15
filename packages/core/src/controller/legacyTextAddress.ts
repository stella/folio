import type { Node as PMNode } from "prosemirror-model";

import { OP_STORIES, type TextPosition } from "@stll/docx-core/ops";

/**
 * The body text position an editor position names in a document that is not
 * a canonical projection: the enclosing paragraph's `paraId` and the offset
 * counted over the paragraph's text, each inline leaf that is not text
 * counting as one unit. Null outside a paragraph with an id.
 */
export const legacyTextAddressAt = (
  doc: PMNode,
  position: number,
): TextPosition | null => {
  if (position < 0 || position > doc.content.size) return null;
  const resolved = doc.resolve(position);
  const paragraph = resolved.parent;
  const paraId: unknown = paragraph.attrs["paraId"];
  if (!paragraph.isTextblock || typeof paraId !== "string" || paraId === "")
    return null;
  const target = resolved.parentOffset;
  let offset = 0;
  paragraph.descendants((node, at) => {
    if (at >= target) return false;
    if (!node.isLeaf) return true;
    offset += node.isText ? Math.min(node.text?.length ?? 0, target - at) : 1;
    return false;
  });
  return { story: OP_STORIES.MAIN, blockId: paraId, offset };
};
