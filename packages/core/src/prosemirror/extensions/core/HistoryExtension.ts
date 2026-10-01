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

const closeGroup = (view: EditorView): false => {
  view.dispatch(closeHistory(view.state.tr));
  return false;
};

/**
 * A paste or drop never joins the edit before it. prosemirror-history
 * otherwise merges it with an adjacent edit made within `newGroupDelay`, and
 * that merge depends on the mode: a deletion removes text in editing mode
 * (adjacent, so merged) but only marks it in suggesting mode (no mapped range,
 * so not), and one undo would then revert different edits in the two modes.
 */
const pasteBoundaryPlugin = (): Plugin =>
  new Plugin({
    props: {
      handleDOMEvents: {
        paste: closeGroup,
        drop: closeGroup,
      },
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
        pasteBoundaryPlugin(),
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
