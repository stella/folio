import type { ParagraphBlock } from "./types";

export const JUSTIFIED_LIST_FINAL_LINE_MAX_SHRINK_RATIO = 0.025;
export const JUSTIFIED_LIST_SPACE_CONTRACTION_RATIO = 0.32;

export const supportsJustifiedListFinalLineContraction = (block: ParagraphBlock): boolean =>
  block.attrs?.justificationCompatibility?.type !== "legacy" &&
  block.attrs?.listMarker !== undefined;
