import type { ParagraphPropsPatch, RunPropsPatch } from "@stll/docx-core/ops";
import type { Node as PMNode } from "prosemirror-model";
import type { Command, EditorState } from "prosemirror-state";

/** Command meaning before canonical positions are resolved by the session. */
export type CanonicalCommandIntent =
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
  [{ type: "formatRun", from: state.selection.from, to: state.selection.to, patch }] as const;

export const withCanonicalParagraphFormatting = (
  command: Command,
  patch: ParagraphPropsPatch | ((paragraph: PMNode) => ParagraphPropsPatch),
): Command =>
  withCanonicalCommand(command, (state) => {
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
  });
