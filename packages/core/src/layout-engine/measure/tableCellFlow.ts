import { measuredLineRangeHeight } from "../lineFlow";
import {
  collapseParagraphSpacing,
  getParagraphSpacingAfter,
  getParagraphSpacingBefore,
  isEmptyParagraph,
} from "../paragraphSpacing";
import type { FlowBlock, Measure, ParagraphBlock, ParagraphMeasure } from "../types";
import { isFloatingTextBoxBlock } from "../types";

export type TableCellBlockPlacement = {
  top: number;
  contentTop: number;
  contentHeight: number;
  leadingSpacing: number;
};

export type TableCellFlowState = {
  height: number;
  previousParagraphWasEmpty: boolean;
  trailingSpacing: number;
};

export const createTableCellFlowState = (): TableCellFlowState => ({
  height: 0,
  previousParagraphWasEmpty: false,
  trailingSpacing: 0,
});

export const finishTableCellFlow = (state: TableCellFlowState): number =>
  state.height + state.trailingSpacing;

const isSuppressedParagraphMeasure = (measure: ParagraphMeasure): boolean =>
  measure.totalHeight === 0 && measure.lines.every(({ lineHeight }) => lineHeight === 0);

const paragraphSpacing = (
  block: ParagraphBlock,
  measure: ParagraphMeasure,
): { before: number; after: number } => {
  if (isSuppressedParagraphMeasure(measure)) {
    return { before: 0, after: 0 };
  }
  return {
    before: getParagraphSpacingBefore(block),
    after: getParagraphSpacingAfter(block),
  };
};

export const getTableCellBlockContentHeight = (block: FlowBlock, measure: Measure): number => {
  if (block.kind === "textBox" && isFloatingTextBoxBlock(block)) {
    return 0;
  }
  if (block.kind === "paragraph" && measure.kind === "paragraph") {
    return measuredLineRangeHeight(measure.lines, 0, measure.lines.length);
  }
  if ("totalHeight" in measure) {
    return measure.totalHeight;
  }
  if ("height" in measure) {
    return measure.height;
  }
  return 0;
};

/**
 * Place one block in an unpaginated table-cell story.
 *
 * Adjacent paragraph spacing collapses to the larger side. Consecutive
 * authored empty paragraphs remain independent vertical spacers, including
 * at the end of the cell.
 */
export const placeTableCellBlock = (
  state: TableCellFlowState,
  block: FlowBlock,
  measure: Measure,
): TableCellBlockPlacement => {
  const top = state.height;
  const contentHeight = getTableCellBlockContentHeight(block, measure);
  if (
    block.kind === "paragraph" &&
    measure.kind === "paragraph" &&
    isSuppressedParagraphMeasure(measure)
  ) {
    return { top, contentTop: top, contentHeight: 0, leadingSpacing: 0 };
  }
  if (block.kind !== "paragraph" || measure.kind !== "paragraph") {
    const leadingSpacing = state.trailingSpacing;
    const contentTop = top + leadingSpacing;
    state.height = contentTop + contentHeight;
    state.previousParagraphWasEmpty = false;
    state.trailingSpacing = 0;
    return { top, contentTop, contentHeight, leadingSpacing };
  }

  const spacing = paragraphSpacing(block, measure);
  const empty = isEmptyParagraph(block);
  const leadingSpacing =
    empty && state.previousParagraphWasEmpty
      ? spacing.before + state.trailingSpacing
      : collapseParagraphSpacing({ before: spacing.before, after: state.trailingSpacing });
  const contentTop = top + leadingSpacing;
  state.height = contentTop + contentHeight;
  state.previousParagraphWasEmpty = empty;
  state.trailingSpacing = spacing.after;
  return { top, contentTop, contentHeight, leadingSpacing };
};
