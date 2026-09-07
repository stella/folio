import type { ParagraphBlock } from "./types";

export const JUSTIFIED_FINAL_LINE_MAX_SHRINK_RATIO = 0.025;
export const JUSTIFIED_FINAL_LINE_SPACE_CONTRACTION_RATIO = 0.32;

export const supportsJustifiedFinalLineContraction = (block: ParagraphBlock): boolean =>
  block.attrs?.alignment === "justify" &&
  block.attrs.justificationCompatibility?.type !== "legacy" &&
  (block.attrs.listMarker !== undefined || (block.attrs.indent?.left ?? 0) > 0);
