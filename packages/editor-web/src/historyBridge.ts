/**
 * Mirrors the canonical session's undo history to a host that keeps its own
 * undo stack (VS Code's `CustomDocumentEditEvent`s). The host adds one entry
 * per undo step the session opens and sends undo / redo back one step at a
 * time, so both stacks stay the same length.
 */

import { Plugin, PluginKey } from "prosemirror-state";
import type { EditorState } from "prosemirror-state";

import { canonicalHistoryEffect } from "@stll/folio-core/controller/canonicalSession";
import type { CanonicalHistoryEffect } from "@stll/folio-core/controller/canonicalSession";

/**
 * What a document change did to the undo history:
 *  - `newStep`: opened a new undo step (typing after a pause, Enter, a paste).
 *  - `sameStep`: joined the step the previous change opened, or stayed out of
 *    history altogether.
 *  - `undoRedo`: was an undo or a redo.
 */
export type HistoryChange = "newStep" | "sameStep" | "undoRedo";

const HISTORY_CHANGES = {
  opened: "newStep",
  joined: "sameStep",
  undo: "undoRedo",
  redo: "undoRedo",
} as const satisfies Record<CanonicalHistoryEffect, HistoryChange>;

type BridgeState = {
  /**
   * One token per editor state lineage. A document load builds a fresh state,
   * and so a fresh token: its history starting over is not an edit.
   */
  readonly lineage: object;
  /** What the session said the last document change does to its history. */
  readonly effect: CanonicalHistoryEffect | null;
};

const historyBridgeKey = new PluginKey<BridgeState>("folioHistoryBridge");

const classify = (
  state: EditorState,
  previous: EditorState,
): HistoryChange | null => {
  const current = historyBridgeKey.getState(state);
  const before = historyBridgeKey.getState(previous);
  if (
    current === undefined ||
    before === undefined ||
    current.lineage !== before.lineage ||
    state.doc === previous.doc
  ) {
    return null;
  }
  // A change the session did not stage stays out of its history.
  return current.effect === null ? "sameStep" : HISTORY_CHANGES[current.effect];
};

/**
 * Body-editor plugin (pass it through `DocxEditor`'s `plugins`) that reports
 * each canonical document change and whether it opened a new undo step. Keep
 * one instance per mounted editor.
 */
export const createHistoryBridgePlugin = (
  onChange: (change: HistoryChange) => void,
): Plugin =>
  new Plugin<BridgeState>({
    key: historyBridgeKey,
    state: {
      init: () => ({ lineage: {}, effect: null }),
      apply: (transaction, value) => {
        if (!transaction.docChanged) {
          return value;
        }
        const effect = canonicalHistoryEffect(transaction);
        return effect === value.effect
          ? value
          : { lineage: value.lineage, effect };
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
