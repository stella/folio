import { panic } from "better-result";

import type { FlowBlock, TextRun } from "../layout-engine/types";
import type { Paragraph, StyleDefinitions, TextFormatting } from "../types/document";
import { STYLE_TOGGLE_KEYS } from "../utils/textFormattingMerge";

export const STYLE_TOGGLE_TEXT = "Style toggle cascade";

const ACTIVE_STYLE_TOGGLES = {
  allCaps: true,
  bold: true,
  boldCs: true,
  emboss: true,
  hidden: true,
  imprint: true,
  italic: true,
  italicCs: true,
  outline: true,
  shadow: true,
  smallCaps: true,
  strike: true,
} as const satisfies TextFormatting & Record<(typeof STYLE_TOGGLE_KEYS)[number], true>;

export const EXPECTED_CANCELLED_STYLE_TOGGLES = {
  allCaps: false,
  bold: false,
  complexScriptBold: false,
  complexScriptItalic: false,
  emboss: false,
  hidden: false,
  imprint: false,
  italic: false,
  smallCaps: false,
  strike: false,
  textOutline: false,
  textShadow: false,
} as const;

export const STYLE_TOGGLE_FLOW_PROPERTIES = {
  allCaps: "allCaps",
  bold: "bold",
  boldCs: "complexScriptBold",
  emboss: "emboss",
  hidden: "hidden",
  imprint: "imprint",
  italic: "italic",
  italicCs: "complexScriptItalic",
  outline: "textOutline",
  shadow: "textShadow",
  smallCaps: "smallCaps",
  strike: "strike",
} as const satisfies Record<
  (typeof STYLE_TOGGLE_KEYS)[number],
  keyof typeof EXPECTED_CANCELLED_STYLE_TOGGLES
>;

export const makeStyleToggleDefinitions = (): StyleDefinitions => ({
  styles: [
    {
      styleId: "ToggleParagraph",
      type: "paragraph",
      name: "Toggle Paragraph",
      rPr: { ...ACTIVE_STYLE_TOGGLES },
    },
    {
      styleId: "ToggleCharacter",
      type: "character",
      name: "Toggle Character",
      rPr: { ...ACTIVE_STYLE_TOGGLES },
    },
  ],
});

export const makeStyleToggleParagraph = (): Paragraph => ({
  type: "paragraph",
  formatting: { styleId: "ToggleParagraph" },
  content: [
    {
      type: "run",
      formatting: { styleId: "ToggleCharacter" },
      content: [{ type: "text", text: STYLE_TOGGLE_TEXT }],
    },
  ],
});

export const findStyleToggleRun = (blocks: readonly FlowBlock[]): TextRun => {
  for (const block of blocks) {
    if (block.kind !== "paragraph") {
      continue;
    }
    const run = block.runs.find(
      (candidate): candidate is TextRun =>
        candidate.kind === "text" && candidate.text === STYLE_TOGGLE_TEXT,
    );
    if (run) {
      return run;
    }
  }
  return panic(`Expected text run: ${STYLE_TOGGLE_TEXT}`);
};
