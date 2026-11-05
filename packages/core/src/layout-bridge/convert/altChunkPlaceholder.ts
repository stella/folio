/**
 * What a page shows for a `w:altChunk`, whose content Folio cannot import.
 *
 * The node stays a preserved leaf, so saving replays it untouched. These blocks
 * exist only in layout: every position in them is the boundary after the node,
 * so a click lands outside the chunk, in the nearest text.
 */

import type { Node as PMNode } from "prosemirror-model";

import { isAltChunkMarkup } from "../../docx/altChunk";
import type { ParagraphBlock } from "../../layout-engine/types";
import { expectPreservedBlockAttrs } from "../../prosemirror/attrs";
import { detectBaseDirection } from "../../utils/baseDirection";
import {
  ALT_CHUNK_CONTENT,
  type FlowConversionOptions,
} from "./flowConversionShared";
import { convertParagraph } from "./paragraphConversion";

/** The marker where no editor supplies a localised one. */
export const ALT_CHUNK_DEFAULT_MARKER_TEXT = "Content not displayed";

/** Tints the block, so it does not read as part of the editable text. */
const PLACEHOLDER_SHADING = "#F2F2F2";
const MARKER_COLOR = "#595959";

/** Past this, further lines of a chunk's text run together in the last block. */
const MAX_PLACEHOLDER_PARAGRAPHS = 2_000;

const altChunkNodes = new WeakMap<PMNode, boolean>();

const isAltChunkNode = (node: PMNode): boolean => {
  const known = altChunkNodes.get(node);
  if (known !== undefined) {
    return known;
  }
  const isAltChunk = isAltChunkMarkup(expectPreservedBlockAttrs(node).xml);
  altChunkNodes.set(node, isAltChunk);
  return isAltChunk;
};

const placeholderLines = (text: string): string[] => {
  const lines = text.split(/\r\n|\r|\n/u);
  if (lines.length <= MAX_PLACEHOLDER_PARAGRAPHS) {
    return lines;
  }
  const kept = lines.slice(0, MAX_PLACEHOLDER_PARAGRAPHS - 1);
  kept.push(lines.slice(MAX_PLACEHOLDER_PARAGRAPHS - 1).join(" "));
  return kept;
};

/**
 * The placeholder for a preserved block at `pos`, or nothing when the block is
 * not an altChunk. Only the first paragraph addresses the node.
 */
export const convertAltChunkPlaceholder = (
  node: PMNode,
  pos: number,
  options: FlowConversionOptions,
): ParagraphBlock[] => {
  if (node.type.name !== "preservedBlock" || !isAltChunkNode(node)) {
    return [];
  }
  const { readerText } = expectPreservedBlockAttrs(node);
  const hasText =
    options.altChunkContent !== ALT_CHUNK_CONTENT.markerOnly &&
    typeof readerText === "string" &&
    readerText.trim() !== "";
  const isMarker = !hasText;
  const text = hasText
    ? readerText
    : (options.altChunkMarkerText ?? ALT_CHUNK_DEFAULT_MARKER_TEXT);
  const { schema } = node.type;
  const paragraphType = schema.nodes["paragraph"];
  if (!paragraphType) {
    return [];
  }
  const after = pos + node.nodeSize;
  const blocks: ParagraphBlock[] = [];
  for (const [index, line] of placeholderLines(text).entries()) {
    const block = convertParagraph(
      paragraphType.create(
        detectBaseDirection(line) === "rtl"
          ? { direction: { source: "auto" } }
          : null,
        line.length > 0 ? schema.text(line) : null,
      ),
      pos,
      options,
    );
    delete block.pmStart;
    delete block.pmEnd;
    if (index === 0) {
      block.pmStart = pos;
      block.pmEnd = after;
    }
    block.fixedPosition = after;
    block.attrs = { ...block.attrs, shading: PLACEHOLDER_SHADING };
    for (const run of block.runs) {
      if (run.kind !== "text") {
        continue;
      }
      run.pmStart = after;
      run.pmEnd = after;
      if (isMarker) {
        run.italic = true;
        run.color = MARKER_COLOR;
      }
    }
    blocks.push(block);
  }
  return blocks;
};
