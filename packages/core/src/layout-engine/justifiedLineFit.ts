import type { ParagraphBlock } from "./types";

export const JUSTIFIED_LIST_FINAL_LINE_MAX_SHRINK_RATIO = 0.025;
export const JUSTIFIED_LIST_SPACE_CONTRACTION_RATIO = 0.32;

type JustifiedListFinalLineShrinkBudgetOptions = {
  availableWidth: number;
  compressibleSpaceWidth: number;
};

export const calculateJustifiedListFinalLineShrinkBudget = ({
  availableWidth,
  compressibleSpaceWidth,
}: JustifiedListFinalLineShrinkBudgetOptions): number =>
  Math.min(
    Math.max(0, availableWidth) * JUSTIFIED_LIST_FINAL_LINE_MAX_SHRINK_RATIO,
    Math.max(0, compressibleSpaceWidth) * JUSTIFIED_LIST_SPACE_CONTRACTION_RATIO,
  );

export const supportsJustifiedListFinalLineContraction = (block: ParagraphBlock): boolean =>
  block.attrs?.justificationCompatibility?.type !== "legacy" &&
  block.attrs?.listMarker !== undefined;
