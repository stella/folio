import type { ParagraphPropsPatch, RunPropsPatch } from "@stll/docx-core/ops";
import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";
import { canonicalSelectionRange } from "./canonicalSelectionRange";

/** Command meaning before canonical positions are resolved by the session. */
export type CanonicalCommandIntent =
  | { type: "generateTOC"; at: number; title: string }
  | { type: "setHyperlink"; from: number; to: number; href: string; tooltip?: string }
  | { type: "removeHyperlink"; from: number; to: number; hyperlinkStyleId?: string }
  | {
      type: "insertHyperlink";
      from: number;
      to: number;
      text: string;
      href: string;
      tooltip?: string;
    }
  | { type: "formatRun"; from: number; to: number; patch: RunPropsPatch }
  | { type: "formatParagraph"; at: number; patch: ParagraphPropsPatch }
  | { type: "toggleList"; kind: "bullet" | "decimal" }
  | { type: "changeListLevel"; direction: "increase" | "decrease" }
  | { type: "removeList" }
  | { type: "restartNumbering"; start?: number }
  | { type: "continueNumbering" }
  | { type: "insertBreak"; from: number; to: number; breakType: "page" | "textWrapping" };

type CanonicalCommandDescriptor = (state: EditorState) => readonly CanonicalCommandIntent[];
const descriptors = new WeakMap<Command, CanonicalCommandDescriptor>();

/** Registration leaves execution and side-effect-free command probes unchanged. */
export const withCanonicalCommand = (
  command: Command,
  descriptor: CanonicalCommandDescriptor,
): Command => {
  descriptors.set(command, descriptor);
  return command;
};

export const getCanonicalCommandIntents = (command: Command, state: EditorState) =>
  descriptors.get(command)?.(state);

export const canonicalRunFormatting = (state: EditorState, patch: RunPropsPatch) =>
  [{ type: "formatRun", ...canonicalSelectionRange(state), patch }] as const;

const canonicalParagraphFormatting = (
  state: EditorState,
  patch: ParagraphPropsPatch | ((paragraph: PMNode) => ParagraphPropsPatch),
) => {
  const intents: CanonicalCommandIntent[] = [];
  state.doc.nodesBetween(state.selection.from, state.selection.to, (node, position) => {
    if (node.type.name !== "paragraph") return;
    intents.push({
      type: "formatParagraph",
      at: position + 1,
      patch: typeof patch === "function" ? patch(node) : patch,
    });
  });
  return intents;
};

export const withCanonicalParagraphFormatting = (
  command: Command,
  patch: ParagraphPropsPatch | ((paragraph: PMNode) => ParagraphPropsPatch),
): Command => withCanonicalCommand(command, (state) => canonicalParagraphFormatting(state, patch));

/** Commands that read the selection start apply that same patch across selected paragraphs. */
export const withCanonicalStartParagraphFormatting = (
  command: Command,
  patch: (paragraph: PMNode) => ParagraphPropsPatch,
): Command =>
  withCanonicalCommand(command, (state) => {
    const paragraph = state.selection.$from.parent;
    if (paragraph.type.name !== "paragraph") return [];
    return canonicalParagraphFormatting(state, patch(paragraph));
  });
