/**
 * Mirrors the body editor's undo history to a host that keeps its own undo
 * stack (VS Code's `CustomDocumentEditEvent`s). The host adds one entry per
 * undo step the editor opens and sends undo / redo back one step at a time, so
 * both stacks stay the same length.
 */

import { isHistoryTransaction, undoDepth } from "prosemirror-history";
import { Plugin, PluginKey, Transaction } from "prosemirror-state";
import type { EditorState } from "prosemirror-state";

/**
 * What a document change did to the undo history:
 *  - `newStep`: opened a new undo step (typing after a pause, Enter, a paste).
 *  - `sameStep`: joined the step the previous change opened, or stayed out of
 *    history altogether.
 *  - `undoRedo`: was an undo or a redo.
 */
export type HistoryChange = "newStep" | "sameStep" | "undoRedo";

type BridgeState = {
  /**
   * One token per editor state lineage. A document load builds a fresh state,
   * and so a fresh token: its history starting over is not an edit.
   */
  readonly lineage: object;
  readonly lastWasUndoRedo: boolean;
};

const historyBridgeKey = new PluginKey<BridgeState>("folioHistoryBridge");

/** A transaction appended by another plugin belongs to the one that caused it. */
const rootTransaction = (transaction: Transaction): Transaction => {
  const root: unknown = transaction.getMeta("appendedTransaction");
  return root instanceof Transaction ? root : transaction;
};

const classify = (state: EditorState, previous: EditorState): HistoryChange | null => {
  const current = historyBridgeKey.getState(state);
  const before = historyBridgeKey.getState(previous);
  if (current === undefined || before === undefined || current.lineage !== before.lineage) {
    return null;
  }
  if (state.doc === previous.doc) {
    return null;
  }
  if (current.lastWasUndoRedo) {
    return "undoRedo";
  }
  // prosemirror-history's event count changes exactly when a change opens a
  // new event: it grows by one, or drops back to the depth limit when the
  // oldest events are trimmed. A change folded into the open event keeps it.
  return undoDepth(state) === undoDepth(previous) ? "sameStep" : "newStep";
};

/**
 * Body-editor plugin (pass it through `DocxEditor`'s `plugins`) that reports
 * each document change and whether it opened a new undo step. Keep one
 * instance per mounted editor.
 */
export const createHistoryBridgePlugin = (onChange: (change: HistoryChange) => void): Plugin =>
  new Plugin<BridgeState>({
    key: historyBridgeKey,
    state: {
      init: () => ({ lineage: {}, lastWasUndoRedo: false }),
      apply: (transaction, value) => {
        if (!transaction.docChanged) {
          return value;
        }
        const lastWasUndoRedo = isHistoryTransaction(rootTransaction(transaction));
        return lastWasUndoRedo === value.lastWasUndoRedo
          ? value
          : { lineage: value.lineage, lastWasUndoRedo };
      },
    },
    view: () => ({
      update: (view, previous) => {
        const change = classify(view.state, previous);
        if (change !== null) {
          onChange(change);
        }
      },
    }),
  });
