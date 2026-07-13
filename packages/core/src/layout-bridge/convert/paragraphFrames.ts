import { getParagraphFrame, markParagraphFrameTextBox } from "../../layout-engine/paragraphFrame";
import type { ParagraphFrame } from "../../layout-engine/paragraphFrame";
import { setTextBoxGroupId } from "../../layout-engine/textBoxGroup";
import type {
  BlockId,
  FlowBlock,
  ImageRunPosition,
  ParagraphBlock,
  TextBoxBlock,
} from "../../layout-engine/types";
import { pixelsToEmu } from "../../utils/units";

type BlockIdFactory = () => BlockId;

const FRAME_PROPERTIES = [
  "width",
  "height",
  "hAnchor",
  "vAnchor",
  "x",
  "y",
  "xAlign",
  "yAlign",
  "wrap",
] as const satisfies readonly (keyof ParagraphFrame)[];

function framesEqual(left: ParagraphFrame, right: ParagraphFrame): boolean {
  return FRAME_PROPERTIES.every((property) => left[property] === right[property]);
}

function horizontalRelativeTo(
  anchor: ParagraphFrame["hAnchor"],
): NonNullable<NonNullable<ImageRunPosition["horizontal"]>["relativeTo"]> {
  if (anchor === "page" || anchor === "margin") {
    return anchor;
  }
  return "column";
}

function verticalRelativeTo(
  anchor: ParagraphFrame["vAnchor"],
): NonNullable<NonNullable<ImageRunPosition["vertical"]>["relativeTo"]> {
  if (anchor === "page" || anchor === "margin") {
    return anchor;
  }
  return "paragraph";
}

function framePosition(frame: ParagraphFrame): ImageRunPosition | undefined {
  const horizontal =
    frame.x !== undefined || frame.xAlign !== undefined || frame.hAnchor !== undefined
      ? {
          relativeTo: horizontalRelativeTo(frame.hAnchor),
          ...(frame.x !== undefined ? { posOffset: pixelsToEmu(frame.x) } : {}),
          ...(frame.xAlign !== undefined ? { align: frame.xAlign } : {}),
        }
      : undefined;
  const verticalAlign = frame.yAlign === "inline" ? undefined : frame.yAlign;
  const vertical =
    frame.y !== undefined || verticalAlign !== undefined || frame.vAnchor !== undefined
      ? {
          relativeTo: verticalRelativeTo(frame.vAnchor),
          ...(frame.y !== undefined ? { posOffset: pixelsToEmu(frame.y) } : {}),
          ...(verticalAlign !== undefined ? { align: verticalAlign } : {}),
        }
      : undefined;

  if (horizontal === undefined && vertical === undefined) {
    return undefined;
  }
  return {
    ...(horizontal !== undefined ? { horizontal } : {}),
    ...(vertical !== undefined ? { vertical } : {}),
  };
}

function frameWrapType(frame: ParagraphFrame): NonNullable<TextBoxBlock["wrapType"]> {
  switch (frame.wrap) {
    case "notBeside":
      return "topAndBottom";
    case "none":
      return "inFront";
    case "through":
      return "through";
    case "tight":
      return "tight";
    case "around":
    case "auto":
    case undefined:
      return "square";
    default:
      frame.wrap satisfies never;
      return "square";
  }
}

function toFrameTextBox(
  content: ParagraphBlock[],
  frame: ParagraphFrame,
  nextBlockId: BlockIdFactory,
): TextBoxBlock {
  const first = content.at(0);
  const last = content.at(-1);
  const textBox: TextBoxBlock = {
    kind: "textBox",
    id: nextBlockId(),
    width: frame.width ?? 200,
    margins: { top: 0, right: 0, bottom: 0, left: 0 },
    content,
    displayMode: "float",
    wrapType: frameWrapType(frame),
    wrapText: "bothSides",
    ...(first?.pmStart !== undefined ? { pmStart: first.pmStart } : {}),
    ...(last?.pmEnd !== undefined ? { pmEnd: last.pmEnd } : {}),
  };
  if (frame.height !== undefined) {
    textBox.height = frame.height;
  }
  const position = framePosition(frame);
  if (position !== undefined) {
    textBox.position = position;
  }
  markParagraphFrameTextBox(textBox);
  return textBox;
}

export function groupParagraphFrames(
  blocks: FlowBlock[],
  nextBlockId: BlockIdFactory,
): FlowBlock[] {
  const grouped: FlowBlock[] = [];
  let index = 0;
  let activeFrameSetId: string | undefined;

  while (index < blocks.length) {
    const block = blocks[index];
    if (block?.kind !== "paragraph") {
      if (block) {
        grouped.push(block);
      }
      activeFrameSetId = undefined;
      index += 1;
      continue;
    }

    const frame = getParagraphFrame(block);
    if (frame === undefined) {
      grouped.push(block);
      if (
        block.runs.length > 0 ||
        activeFrameSetId === undefined ||
        !hasFollowingParagraphFrame(blocks, index + 1)
      ) {
        activeFrameSetId = undefined;
      }
      index += 1;
      continue;
    }

    activeFrameSetId ??= `paragraph-frame-${block.pmStart ?? block.id}`;
    const content: ParagraphBlock[] = [];
    while (index < blocks.length) {
      const paragraph = blocks[index];
      if (paragraph?.kind !== "paragraph") {
        break;
      }
      const paragraphFrame = getParagraphFrame(paragraph);
      if (paragraphFrame === undefined || !framesEqual(frame, paragraphFrame)) {
        break;
      }
      content.push(paragraph);
      index += 1;
    }

    const textBox = toFrameTextBox(content, frame, nextBlockId);
    setTextBoxGroupId(textBox, activeFrameSetId);
    grouped.push(textBox);
  }

  return grouped;
}

function hasFollowingParagraphFrame(blocks: FlowBlock[], startIndex: number): boolean {
  for (let index = startIndex; index < blocks.length; index++) {
    const block = blocks[index];
    if (block?.kind !== "paragraph") {
      return false;
    }
    if (getParagraphFrame(block) !== undefined) {
      return true;
    }
    if (block.runs.length > 0) {
      return false;
    }
  }
  return false;
}
