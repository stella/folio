import { panic } from "better-result";

import { measuredLineRangeHeight } from "../lineFlow";
import {
  collapseParagraphSpacing,
  getParagraphSpacingAfter,
  getParagraphSpacingBefore,
} from "../paragraphSpacing";
import type { ParagraphBlock, ParagraphMeasure } from "../types";

export type TextBoxParagraphPlacement = {
  leadingSpacing: number;
  contentHeight: number;
};

export type TextBoxParagraphLayout = {
  placements: TextBoxParagraphPlacement[];
  totalHeight: number;
};

/**
 * Lay out a text box's paragraph story as one unpaginated flow. Adjacent
 * before/after spacing collapses to the larger value, matching body flow.
 */
export function layoutTextBoxParagraphs(
  blocks: readonly ParagraphBlock[],
  measures: readonly ParagraphMeasure[],
): TextBoxParagraphLayout {
  if (blocks.length !== measures.length) {
    panic("layoutTextBoxParagraphs: block and measure counts must match");
  }

  const placements: TextBoxParagraphPlacement[] = [];
  let totalHeight = 0;
  let trailingSpacing = 0;

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]!; // SAFETY: index < blocks.length
    const measure = measures[index]!; // SAFETY: equal lengths checked above

    const leadingSpacing = collapseParagraphSpacing({
      before: getParagraphSpacingBefore(block),
      after: trailingSpacing,
    });
    const contentHeight = measuredLineRangeHeight(measure.lines, 0, measure.lines.length);
    placements.push({ leadingSpacing, contentHeight });
    totalHeight += leadingSpacing + contentHeight;
    trailingSpacing = getParagraphSpacingAfter(block);
  }

  return { placements, totalHeight: totalHeight + trailingSpacing };
}
