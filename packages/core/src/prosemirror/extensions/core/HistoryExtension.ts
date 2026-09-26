/**
 * History Extension — undo/redo via prosemirror-history
 */

import { history, undo, redo } from "prosemirror-history";

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
