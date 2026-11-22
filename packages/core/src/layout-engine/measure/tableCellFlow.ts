import { measuredLineRangeHeight } from "../lineFlow";
import {
  collapseParagraphSpacing,
  getParagraphSpacingAfter,
  getParagraphSpacingBefore,
  isEmptyParagraph,
} from "../paragraphSpacing";
import type {
  FlowBlock,
  Measure,
  ParagraphBlock,
  ParagraphMeasure,
} from "../types";
import { isFloatingTextBoxBlock } from "../types";

export type TableCellBlockPlacement = {
  top: number;
  contentTop: number;
  contentHeight: number;
  leadingSpacing: number;
};

const TABLE_CELL_FLOW_POSITION = {
  start: "start",
  interior: "interior",
} as const;
type TableCellFlowPosition =
  (typeof TABLE_CELL_FLOW_POSITION)[keyof typeof TABLE_CELL_FLOW_POSITION];

const TABLE_CELL_TRAILING_SPACING_TYPE = {
  none: "none",
  retained: "retained",
  suppressedAtCellEnd: "suppressedAtCellEnd",
} as const;

type TableCellTrailingSpacing =
  | { type: typeof TABLE_CELL_TRAILING_SPACING_TYPE.none }
  | { type: typeof TABLE_CELL_TRAILING_SPACING_TYPE.retained; value: number }
  | {
      type: typeof TABLE_CELL_TRAILING_SPACING_TYPE.suppressedAtCellEnd;
      value: number;
    };

export type TableCellFlowState = {
  height: number;
  position: TableCellFlowPosition;
  previousParagraphWasEmpty: boolean;
  trailingSpacing: TableCellTrailingSpacing;
};

export const createTableCellFlowState = (): TableCellFlowState => ({
  height: 0,
  position: TABLE_CELL_FLOW_POSITION.start,
  previousParagraphWasEmpty: false,
  trailingSpacing: { type: TABLE_CELL_TRAILING_SPACING_TYPE.none },
});

const trailingSpacingValue = (spacing: TableCellTrailingSpacing): number => {
  switch (spacing.type) {
    case TABLE_CELL_TRAILING_SPACING_TYPE.none:
      return 0;
    case TABLE_CELL_TRAILING_SPACING_TYPE.retained:
    case TABLE_CELL_TRAILING_SPACING_TYPE.suppressedAtCellEnd:
      return spacing.value;
    default: {
      const unhandled: never = spacing;
      return unhandled;
    }
  }
};

export const finishTableCellFlow = (state: TableCellFlowState): number => {
  if (
    state.trailingSpacing.type ===
    TABLE_CELL_TRAILING_SPACING_TYPE.suppressedAtCellEnd
  ) {
    return state.height;
  }
  return state.height + trailingSpacingValue(state.trailingSpacing);
};

const isSuppressedParagraphMeasure = (measure: ParagraphMeasure): boolean =>
  measure.totalHeight === 0 &&
  measure.lines.every(({ lineHeight }) => lineHeight === 0);

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

export const getTableCellBlockContentHeight = (
  block: FlowBlock,
  measure: Measure,
): number => {
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
  // Floating text boxes paint outside the cell's block flow. Keeping this a
  // no-op matches the painter and preserves the boundary for visible content.
  if (block.kind === "textBox" && isFloatingTextBoxBlock(block)) {
    return { top, contentTop: top, contentHeight: 0, leadingSpacing: 0 };
  }
  if (block.kind !== "paragraph" || measure.kind !== "paragraph") {
    const leadingSpacing = trailingSpacingValue(state.trailingSpacing);
    const contentTop = top + leadingSpacing;
    state.height = contentTop + contentHeight;
    state.position = TABLE_CELL_FLOW_POSITION.interior;
    state.previousParagraphWasEmpty = false;
    state.trailingSpacing = { type: TABLE_CELL_TRAILING_SPACING_TYPE.none };
    return { top, contentTop, contentHeight, leadingSpacing };
  }

  const spacing = paragraphSpacing(block, measure);
  const empty = isEmptyParagraph(block);
  // The first visible paragraph in the cell drops automatic space before.
  const before =
    state.position === TABLE_CELL_FLOW_POSITION.start &&
    block.attrs?.automaticSpacing?.before === true
      ? 0
      : spacing.before;
  const trailingSpacing = trailingSpacingValue(state.trailingSpacing);
  const leadingSpacing =
    empty && state.previousParagraphWasEmpty
      ? before + trailingSpacing
      : collapseParagraphSpacing({ before, after: trailingSpacing });
  const contentTop = top + leadingSpacing;
  state.height = contentTop + contentHeight;
  state.position = TABLE_CELL_FLOW_POSITION.interior;
  state.previousParagraphWasEmpty = empty;
  // Numbered lists keep their automatic trailing margin at the cell edge.
  const suppressAfterAtCellEnd =
    block.attrs?.automaticSpacing?.after === true &&
    block.attrs.numPr?.kind !== "reference";
  state.trailingSpacing = {
    type: suppressAfterAtCellEnd
      ? TABLE_CELL_TRAILING_SPACING_TYPE.suppressedAtCellEnd
      : TABLE_CELL_TRAILING_SPACING_TYPE.retained,
    value: spacing.after,
  };
  return { top, contentTop, contentHeight, leadingSpacing };
};
