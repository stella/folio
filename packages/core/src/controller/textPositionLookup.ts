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
export const legacyLeafUnits = (node: PMNode): number => {
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
 * position's offset counted in the engine's logical units (see `legacyLeafUnits`),
 * so zero-width anchors before the point do not move it. Null for another
 * story, an unknown paragraph, an offset past its end, or an ambiguous control edge.
 */
export const legacyEditorPositionAt = (
  doc: PMNode,
  position: TextPosition,
): number | null => {
  if (position.story !== OP_STORIES.MAIN) return null;
  const paragraph = findParagraphByParaId(doc, position.blockId);
  return paragraph === null
    ? null
    : legacyParagraphPositionAt({
        paragraph: paragraph.node,
        contentStart: paragraph.from + 1,
        position,
      });
};

type LegacyParagraphPositionOptions = {
  paragraph: PMNode;
  contentStart: number;
  position: TextPosition;
};

/** Resolve within an already located paragraph, without rescanning the document. */
export const legacyParagraphPositionAt = ({
  paragraph,
  contentStart,
  position,
}: LegacyParagraphPositionOptions): number | null => {
  if (!Number.isInteger(position.offset) || position.offset < 0) return null;
  const ordinal = position.zeroWidthBefore;
  if (ordinal !== undefined && (!Number.isInteger(ordinal) || ordinal < 0))
    return null;
  let logical = 0;
  let passedMarkers = 0;
  let candidateOrdinal = -1;
  const candidates = new Set<number>();
  const record = (at: number) => {
    if (logical !== position.offset) return;
    if (ordinal !== undefined && passedMarkers !== ordinal) return;
    // An unspecified marker gap retains the historical after-markers affinity.
    if (passedMarkers < candidateOrdinal) return;
    if (passedMarkers > candidateOrdinal) candidates.clear();
    candidateOrdinal = passedMarkers;
    candidates.add(at);
  };
  const visit = (parent: PMNode, start: number) => {
    record(start);
    parent.forEach((node, offset) => {
      const at = start + offset;
      record(at);
      if (!node.isLeaf) {
        visit(node, at + 1);
      } else {
        const units = legacyLeafUnits(node);
        if (
          node.isText &&
          logical < position.offset &&
          position.offset < logical + units
        ) {
          const before = logical;
          logical = position.offset;
          passedMarkers = 0;
          record(at + position.offset - before);
          logical = before;
        }
        if (units === 0) passedMarkers += 1;
        else {
          logical += units;
          passedMarkers = 0;
        }
      }
      record(at + node.nodeSize);
    });
    record(start + parent.content.size);
  };
  visit(paragraph, contentStart);
  // Container edges can share the same logical offset and marker ordinal.
  // TextPosition cannot distinguish them; refusing avoids inserting into the
  // wrong control. Supporting them requires an explicit address contract.
  return candidates.size === 1
    ? (candidates.values().next().value ?? null)
    : null;
};
