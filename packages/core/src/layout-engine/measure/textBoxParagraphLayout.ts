import { panic } from "better-result";

import { measuredLineRangeHeight } from "../lineFlow";
import {
  collapseParagraphSpacing,
  getParagraphSpacingAfter,
  getParagraphSpacingBefore,
} from "../paragraphSpacing";
import type { ParagraphBlock, ParagraphMeasure, TableBlock, TableMeasure } from "../types";

export type TextBoxContentPlacement = {
  leadingSpacing: number;
  contentHeight: number;
};

export type TextBoxContentLayout = {
  placements: TextBoxContentPlacement[];
  totalHeight: number;
};

type TextBoxContentBlock = ParagraphBlock | TableBlock;
type TextBoxContentMeasure = ParagraphMeasure | TableMeasure;

/**
 * Lay out a text box's block story as one unpaginated flow. Adjacent paragraph
 * before/after spacing collapses to the larger value, matching body flow.
 */
export function layoutTextBoxContent(
  blocks: readonly TextBoxContentBlock[],
  measures: readonly TextBoxContentMeasure[],
): TextBoxContentLayout {
  if (blocks.length !== measures.length) {
    panic("layoutTextBoxContent: block and measure counts must match");
  }

  const placements: TextBoxContentPlacement[] = [];
  let totalHeight = 0;
  let trailingSpacing = 0;

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index]!; // SAFETY: index < blocks.length
    const measure = measures[index]!; // SAFETY: equal lengths checked above

    if (block.kind === "table") {
      if (measure.kind !== "table") {
        panic("layoutTextBoxContent: block and measure kinds must match");
      }
      const leadingSpacing = trailingSpacing;
      const contentHeight = measure.totalHeight;
      placements.push({ leadingSpacing, contentHeight });
      totalHeight += leadingSpacing + contentHeight;
      trailingSpacing = 0;
      continue;
    }

    if (measure.kind !== "paragraph") {
      panic("layoutTextBoxContent: block and measure kinds must match");
    }

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
