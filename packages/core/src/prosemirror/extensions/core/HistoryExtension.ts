/**
 * History Extension — undo/redo via prosemirror-history
 */

import { closeHistory, history, undo, redo } from "prosemirror-history";
import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { createExtension } from "../create";
import type { ExtensionContext, ExtensionRuntime } from "../types";

/**
 * Who answers the undo and redo keys (Mod-z, Mod-y, Mod-Shift-z).
 *  - `"editor"`: the extension binds them. Default.
 *  - `"host"`: nothing binds them. The host keeps its own undo stack and runs
 *    the `undo` / `redo` commands itself, so one press never undoes twice.
 */
export type HistoryShortcutOwner = "editor" | "host";

type HistoryOptions = {
  depth: number;
  newGroupDelay: number;
  shortcuts: HistoryShortcutOwner;
};

type InputKind = "typing" | "deletion" | "structure" | "composition" | "paste" | "drop";

/**
 * History adjacency is based on step maps, which differ for physical deletions
 * and suggestion marks. Group continuous typing explicitly; each deletion,
 * structural edit, paste, drop, or composition starts its own undo event.
 * Capture DOM input before the shared input router or a keymap handles it.
 */
const inputBoundaryPlugin = (): Plugin =>
  new Plugin({
    view(view: EditorView) {
      let previousKind: InputKind | undefined;
      const startInput = (kind: InputKind) => {
        if (kind !== previousKind || (kind !== "typing" && kind !== "composition")) {
          view.dispatch(closeHistory(view.state.tr));
        }
        previousKind = kind;
      };
      const keydown = (event: KeyboardEvent) => {
        if (event.isComposing) return;
        if (event.key === "Backspace" || event.key === "Delete") {
          startInput("deletion");
          return;
        }
        if (event.key === "Enter") {
          startInput("structure");
          return;
        }
        if (event.key.length === 1 && !event.metaKey && !event.ctrlKey) startInput("typing");
      };
      const beforeinput = (event: InputEvent) => {
        if (event.inputType === "insertCompositionText" || event.isComposing) {
          startInput("composition");
          return;
        }
        if (event.inputType === "insertText" || event.inputType === "insertReplacementText") {
          startInput("typing");
          return;
        }
        if (event.inputType.startsWith("delete")) startInput("deletion");
        else if (event.inputType === "insertParagraph" || event.inputType === "insertLineBreak") {
          startInput("structure");
        }
      };
      const paste = () => startInput("paste");
      const drop = () => startInput("drop");
      const compositionstart = () => {
        view.dispatch(closeHistory(view.state.tr));
        previousKind = "composition";
      };
      view.dom.addEventListener("keydown", keydown, true);
      view.dom.addEventListener("beforeinput", beforeinput, true);
      view.dom.addEventListener("paste", paste, true);
      view.dom.addEventListener("drop", drop, true);
      view.dom.addEventListener("compositionstart", compositionstart, true);
      return {
        destroy() {
          view.dom.removeEventListener("keydown", keydown, true);
          view.dom.removeEventListener("beforeinput", beforeinput, true);
          view.dom.removeEventListener("paste", paste, true);
          view.dom.removeEventListener("drop", drop, true);
          view.dom.removeEventListener("compositionstart", compositionstart, true);
        },
      };
    },
  });

const defaultHistoryOptions: HistoryOptions = {
  depth: 100,
  newGroupDelay: 500,
  shortcuts: "editor",
};

export const HistoryExtension = createExtension({
  name: "history",
  defaultOptions: defaultHistoryOptions,
  onSchemaReady(_ctx: ExtensionContext, options: HistoryOptions): ExtensionRuntime {
    return {
      plugins: [
        history({
          depth: options.depth,
          newGroupDelay: options.newGroupDelay,
        }),
        inputBoundaryPlugin(),
      ],
      commands: {
        undo: () => undo,
        redo: () => redo,
      },
      ...(options.shortcuts === "editor" && {
        keyboardShortcuts: {
          "Mod-z": undo,
          "Mod-y": redo,
          "Mod-Shift-z": redo,
        },
      }),
    };
  },
});
