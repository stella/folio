/**
 * History Extension — undo/redo via prosemirror-history
 */

import { closeHistory, history, isHistoryTransaction, undo, redo } from "prosemirror-history";
import { Plugin } from "prosemirror-state";
import type { Command } from "prosemirror-state";
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

type InputKind =
  | "typing"
  | "deleteBackward"
  | "deleteForward"
  | "deletion"
  | "structure"
  | "composition"
  | "paste"
  | "drop";

/**
 * History adjacency is based on step maps, which differ for physical deletions
 * and suggestion marks. Group continuous typing and same-direction deletions explicitly; each
 * structural edit, paste, drop, or composition starts its own undo event.
 * Capture DOM input before the shared input router or a keymap handles it.
 */
const inputBoundaryPlugin = (newGroupDelay: number): Plugin => {
  let previousKind: InputKind | undefined;
  let deletionGroup = 0;
  let previousInputTime = 0;
  let pendingDeletionGroup: number | null = null;
  return new Plugin({
    filterTransaction(tr) {
      const pendingGroup = pendingDeletionGroup;
      pendingDeletionGroup = null;
      if (isHistoryTransaction(tr) || (!tr.docChanged && tr.selectionSet)) {
        previousKind = undefined;
        return true;
      }
      if (
        tr.docChanged &&
        (previousKind === "deleteBackward" ||
          previousKind === "deleteForward" ||
          previousKind === "deletion")
      ) {
        // History's composition token groups empty-map mark steps too. Negative
        // tokens keep semantic deletion groups separate from native compositions.
        if (pendingGroup !== null) {
          tr.setMeta("composition", pendingGroup);
        } else if (!tr.getMeta("appendedTransaction")) {
          // A command between deletion inputs also ends the input sequence.
          previousKind = undefined;
        }
      }
      return true;
    },
    view(view: EditorView) {
      const startInput = (kind: InputKind) => {
        const now = Date.now();
        if (
          kind !== previousKind ||
          kind === "structure" ||
          kind === "paste" ||
          kind === "drop" ||
          now - previousInputTime > newGroupDelay
        ) {
          deletionGroup--;
          view.dispatch(closeHistory(view.state.tr));
        }
        previousInputTime = now;
        previousKind = kind;
        pendingDeletionGroup =
          kind === "deleteBackward" || kind === "deleteForward" || kind === "deletion"
            ? deletionGroup
            : null;
      };
      const keydown = (event: KeyboardEvent) => {
        if (event.isComposing) return;
        if (event.key === "Backspace" || event.key === "Delete") {
          startInput(event.key === "Backspace" ? "deleteBackward" : "deleteForward");
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
        if (event.inputType.startsWith("delete")) {
          if (event.inputType.endsWith("Backward")) startInput("deleteBackward");
          else if (event.inputType.endsWith("Forward")) startInput("deleteForward");
          else startInput("deletion");
        } else if (event.inputType === "insertParagraph" || event.inputType === "insertLineBreak") {
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
};

// An unavailable editor command must still consume its shortcut. Otherwise a
// shifted redo chord can reach native DOM history, whose IME edits bypass the
// editor's transaction and suggestion history.
const consumeHistoryShortcut =
  (command: Command): Command =>
  (state, dispatch, view) => {
    command(state, dispatch, view);
    return true;
  };

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
        inputBoundaryPlugin(options.newGroupDelay),
      ],
      commands: {
        undo: () => undo,
        redo: () => redo,
      },
      ...(options.shortcuts === "editor" && {
        keyboardShortcuts: {
          "Mod-z": consumeHistoryShortcut(undo),
          "Mod-y": consumeHistoryShortcut(redo),
          "Mod-Shift-z": consumeHistoryShortcut(redo),
        },
      }),
    };
  },
});
