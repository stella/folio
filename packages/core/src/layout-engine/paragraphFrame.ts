import type { ParagraphBlock, TextBoxBlock } from "./types";

export type ParagraphFrame = {
  width?: number;
  height?: number;
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

export function setParagraphFrame(block: ParagraphBlock, frame: ParagraphFrame): void {
  paragraphFrames.set(block, frame);
}

export function getParagraphFrame(block: ParagraphBlock): ParagraphFrame | undefined {
  return paragraphFrames.get(block);
}

export function markParagraphFrameTextBox(block: TextBoxBlock): void {
  paragraphFrameTextBoxes.add(block);
}

export function isParagraphFrameTextBox(block: TextBoxBlock): boolean {
  return paragraphFrameTextBoxes.has(block);
}
