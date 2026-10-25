import type { ParagraphBlock, TextBoxBlock } from "../layout-engine/types";

export type ParagraphFrame = {
  width?: number;
  height?: number;
  hSpace?: number;
  vSpace?: number;
  hAnchor?: "text" | "margin" | "page";
  vAnchor?: "text" | "margin" | "page";
  x?: number;
  y?: number;
  xAlign?: "left" | "center" | "right" | "inside" | "outside";
  yAlign?: "top" | "center" | "bottom" | "inside" | "outside" | "inline";
  wrap?: "around" | "auto" | "none" | "notBeside" | "through" | "tight";
};

const paragraphFrames = new WeakMap<ParagraphBlock, ParagraphFrame>();
const paragraphFrameTextBoxes = new WeakSet<TextBoxBlock>();

export const setParagraphFrame = (
  block: ParagraphBlock,
  frame: ParagraphFrame,
): void => {
  paragraphFrames.set(block, frame);
};

export const getParagraphFrame = (
  block: ParagraphBlock,
): ParagraphFrame | undefined => paragraphFrames.get(block);

export const markParagraphFrameTextBox = (block: TextBoxBlock): void => {
  paragraphFrameTextBoxes.add(block);
};

export const isParagraphFrameTextBox = (block: TextBoxBlock): boolean =>
  paragraphFrameTextBoxes.has(block);
