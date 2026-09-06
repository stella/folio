/**
 * editorShortcuts
 *
 * Framework-agnostic keyboard-shortcut logic for the editor.
 * Extracted from the React `useKeyboardShortcuts` hook.
 *
 * The hook itself carries no state — it wires a document `keydown` listener and
 * routes to host callbacks. This module owns the pieces that are not React: the
 * platform/input predicates, the modifier+key classification, and the
 * whole-table deletion check against ProseMirror state. The adapter keeps only
 * the listener lifecycle and the host-callback dispatch.
 */

import type { EditorState, Transaction } from "prosemirror-state";

import { deleteTable, getTableContext } from "../prosemirror/commands/table";

export function isMacPlatform(): boolean {
  if (typeof navigator === "undefined") {
    return false;
  }
  return /Mac|iPod|iPhone|iPad/u.test(navigator.platform);
}

/**
 * `target` is an input-like element that the user is typing into. We must
 * not intercept Delete/Backspace there — only when focus is in the editor
 * surface (or nowhere at all).
 */
export function isFocusInInputLike(
  target: EventTarget | null,
  editorDom: HTMLElement | null | undefined,
): boolean {
  if (!(target instanceof HTMLElement)) {
    return false;
  }
  if (target.tagName === "INPUT" || target.tagName === "TEXTAREA") {
    return true;
  }
  if (target.isContentEditable && target !== editorDom) {
    return true;
  }
  return false;
}

/** The editor action a keydown maps to, or `none` when it is not a shortcut. */
export type EditorKeydownIntent =
  | { type: "deleteSelectedTable" }
  | { type: "openFind" }
  | { type: "print" }
  | { type: "none" };

export type ClassifyEditorKeydownOptions = {
  /** Whether the platform uses Cmd (Mac) rather than Ctrl as the primary modifier. */
  isMac: boolean;
  /** Whether focus is in a non-editor input/textarea/contenteditable. */
  isInputLike: boolean;
};

/**
 * Map a keydown to an editor intent:
 *  - Cmd/Ctrl+F or Cmd/Ctrl+H → open find
 *  - Cmd/Ctrl+P (no auto-repeat) → custom print
 *  - Delete/Backspace (no modifiers, focus not in an input) → delete selected table
 */
export function classifyEditorKeydown(
  e: KeyboardEvent,
  { isMac, isInputLike }: ClassifyEditorKeydownOptions,
): EditorKeydownIntent {
  const cmdOrCtrl = isMac ? e.metaKey : e.ctrlKey;

  if (
    !cmdOrCtrl &&
    !e.shiftKey &&
    !e.altKey &&
    (e.key === "Delete" || e.key === "Backspace") &&
    !isInputLike
  ) {
    return { type: "deleteSelectedTable" };
  }

  if (cmdOrCtrl && !e.shiftKey && !e.altKey) {
    const key = e.key.toLowerCase();
    if (key === "f" || key === "h") {
      return { type: "openFind" };
    }
    if (key === "p" && !e.repeat) {
      return { type: "print" };
    }
  }

  return { type: "none" };
}

/**
 * Which key presses the editor's page-level shortcuts answer.
 *  - `"document"`: every press on the page.
 *  - `"editor"`: only a press whose target is inside the editor.
 *  - `"none"`: no page-level listener; the host dispatches shortcuts itself.
 */
export type KeyboardShortcutScope = "document" | "editor" | "none";

/**
 * The containment surface a scope root has to offer. Structural rather than
 * `Element` so the predicate carries no DOM dependency and answers the same
 * under a server render and a DOM-less test runtime.
 */
export type ShortcutScopeRoot = { contains(other: unknown): boolean };

export type IsKeydownInShortcutScopeOptions = {
  scope: KeyboardShortcutScope;
  /**
   * Elements a press must land inside under `"editor"`: the editor root plus
   * any surface it owns outside that subtree (a portaled dialog). A nullish
   * entry — an unmounted ref — never matches.
   */
  roots: readonly (ShortcutScopeRoot | null | undefined)[];
};

/** `contains` rejects a non-node argument, so screen the target first. */
function isNode(target: unknown): boolean {
  return typeof target === "object" && target !== null && "nodeType" in target;
}

/**
 * Whether a keydown is one the editor answers under `scope`. The adapters own
 * the listener; this decides, from the event target alone, whether the press
 * belongs to this editor or to whatever else the host has on the page.
 */
export function isKeydownInShortcutScope(
  event: { target: unknown },
  { scope, roots }: IsKeydownInShortcutScopeOptions,
): boolean {
  switch (scope) {
    case "document":
      return true;
    case "none":
      return false;
    case "editor": {
      if (!isNode(event.target)) {
        return false;
      }
      return roots.some((root) => root?.contains(event.target) === true);
    }
    default: {
      const exhaustive: never = scope;
      return exhaustive;
    }
  }
}

/**
 * If the selection is a ProseMirror `CellSelection` covering every cell of its
 * table, delete the table and return `true`. Returns `false` for any other
 * selection so the caller can fall back to a layout-overlay table selection.
 */
export function deleteSelectedTable(
  state: EditorState,
  dispatch: (tr: Transaction) => void,
): boolean {
  const sel = state.selection as { $anchorCell?: unknown; forEachCell?: unknown };
  const isCellSel = "$anchorCell" in sel && typeof sel.forEachCell === "function";
  if (!isCellSel) {
    return false;
  }

  const context = getTableContext(state);
  if (!(context.isInTable && context.table)) {
    return false;
  }

  let totalCells = 0;
  context.table.descendants((node) => {
    if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
      totalCells += 1;
    }
  });

  let selectedCells = 0;
  (sel as { forEachCell: (fn: () => void) => void }).forEachCell(() => {
    selectedCells += 1;
  });

  if (totalCells > 0 && selectedCells >= totalCells) {
    deleteTable(state, dispatch);
    return true;
  }

  return false;
}
