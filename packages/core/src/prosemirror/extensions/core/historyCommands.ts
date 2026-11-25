/**
 * Undo and redo: the commands and the keys that run them, apart from the
 * `prosemirror-history` plugin that keeps the stack.
 */

import { redo, undo } from "prosemirror-history";
import type { Command } from "prosemirror-state";

import { withCanonicalHistory } from "../../canonicalCommands";
import type { KeyboardShortcutMap } from "../types";

/**
 * Who answers the undo and redo keys (Mod-z, Mod-y, Mod-Shift-z).
 *  - `"editor"`: the extension binds them. Default.
 *  - `"host"`: nothing binds them. The host keeps its own undo stack and runs
 *    the `undo` / `redo` commands itself, so one press never undoes twice.
 */
export type HistoryShortcutOwner = "editor" | "host";

export const undoCommand = (): Command => withCanonicalHistory(undo, "undo");

export const redoCommand = (): Command => withCanonicalHistory(redo, "redo");

// An unavailable editor command must still consume its shortcut. Otherwise a
// shifted redo chord can reach native DOM history, whose IME edits bypass the
// editor's transaction and suggestion history.
const consumeHistoryShortcut =
  (command: Command): Command =>
  (state, dispatch, view) => {
    command(state, dispatch, view);
    return true;
  };

export const historyKeyboardShortcuts: KeyboardShortcutMap = {
  "Mod-z": consumeHistoryShortcut(undo),
  "Mod-y": consumeHistoryShortcut(redo),
  "Mod-Shift-z": consumeHistoryShortcut(redo),
};
