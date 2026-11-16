import type { Node as PMNode } from "prosemirror-model";

import { OP_STORIES, type TextPosition } from "@stll/docx-core/ops";

import { legacyLeafUnits } from "./textPositionLookup";

/**
 * The body text position an editor position names in a document that is not
 * a canonical projection: the enclosing paragraph's `paraId` and the offset
 * counted in the same logical units and marker gaps as legacyEditorPositionAt.
 * Null outside a paragraph with an id.
 */
export const legacyTextAddressAt = (
  doc: PMNode,
  position: number,
): TextPosition | null => {
  if (
    !Number.isInteger(position) ||
    position < 0 ||
    position > doc.content.size
  )
    return null;
  const resolved = doc.resolve(position);
  let depth = resolved.depth;
  while (depth > 0 && !resolved.node(depth).isTextblock) depth -= 1;
  const paragraph = resolved.node(depth);
  const paraId: unknown = paragraph.attrs["paraId"];
  if (!paragraph.isTextblock || typeof paraId !== "string" || paraId === "")
    return null;
  const target = position - resolved.start(depth);
  let offset = 0;
  paragraph.descendants((node, at) => {
    if (at >= target) return false;
    if (!node.isLeaf) return true;
    const units = legacyLeafUnits(node);
    offset += node.isText ? Math.min(units, target - at) : units;
    return false;
  });
  let logical = 0;
  let zeroWidthBefore = 0;
  let hasMarkers = false;
  paragraph.descendants((node, at) => {
    if (!node.isLeaf) return true;
    const units = legacyLeafUnits(node);
    if (units === 0 && logical === offset) {
      hasMarkers = true;
      if (at < target) zeroWidthBefore += 1;
    }
    logical += units;
    return false;
  });
  return {
    story: OP_STORIES.MAIN,
    blockId: paraId,
    offset,
    ...(hasMarkers ? { zeroWidthBefore } : {}),
  };
};
