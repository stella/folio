// Reads and writes through the owner, and layout values named like the
// attributes, are not painted-attribute access.
import {
  PAINTED_RANGE_SELECTOR,
  paintedStartSelector,
  readPaintedStart,
  writePaintedRange,
} from "../../packages/core/src/utils/paintedPositions";

declare const element: HTMLElement;
declare const root: ParentNode;
declare const fragment: { pmStart?: number; pmEnd?: number };

export const start = readPaintedStart(element) ?? Number.NaN;
writePaintedRange(element, fragment.pmStart, fragment.pmEnd);
export const spans = root.querySelectorAll(`span${PAINTED_RANGE_SELECTOR}`);
export const anchor = root.querySelector(paintedStartSelector(3));
export const blockId = element.dataset["blockId"];
export const layoutStart = fragment.pmStart;
declare const block: { pmStart?: number };
export const known = "pmStart" in block;
export const { blockId: destructuredId } = element.dataset;
export type PositionField = "pmStart" | "pmEnd";
