import { useEffect, useRef } from "react";
import type { RefObject } from "react";

import {
  classifyEditorKeydown,
  deleteSelectedTable,
  isFocusInInputLike,
  isKeydownInShortcutScope,
  isMacPlatform,
} from "@stll/folio-core/managers/editorShortcuts";
import type { KeyboardShortcutScope } from "@stll/folio-core/managers/editorShortcuts";
import type { PagedEditorRef } from "../../paged-editor/PagedEditor";
import { readFindSelectionSeed } from "../dialogs/findReplaceSelectionSeed";
import type { UseFindReplaceReturn } from "../dialogs/useFindReplace";

export type UseKeyboardShortcutsArgs = {
  pagedEditorRef: RefObject<PagedEditorRef | null>;
  findReplace: UseFindReplaceReturn;
  tableSelection: {
    state: { tableIndex: number | null };
    handleAction: (action: "deleteTable") => void;
  };
  /** Triggered on Cmd/Ctrl+P. */
  onDirectPrint: () => void;
  /** Which presses these shortcuts answer. */
  scope: KeyboardShortcutScope;
  /**
   * Elements a press must land inside under the `"editor"` scope. The editor
   * root covers the find/replace dialog, which renders inside it; pass a
   * further root for any surface portaled out of that subtree.
   */
  roots: readonly RefObject<HTMLElement | null>[];
};

/**
 * Page-level keyboard shortcuts:
 *  - Cmd/Ctrl+F → open find dialog with selected text
 *  - Cmd/Ctrl+H → open replace dialog
 *  - Cmd/Ctrl+P → trigger the custom print path (intercepts the OS dialog)
 *  - Delete/Backspace → delete the currently selected table when nothing else
 *    is selected (works with both ProseMirror `CellSelection` whole-table
 *    selections and the layout-overlay table selection). Suppressed when
 *    focus is in a non-editor input/textarea/contenteditable to avoid
 *    deleting tables while the user is typing in a sidebar or dialog.
 *
 * `scope` decides which presses reach that dispatch: `"document"` every press
 * on the page, `"editor"` only presses inside `roots`, `"none"` no listener at
 * all, leaving the shortcuts to the host.
 *
 * Thin React binding: the classification, scope, platform/input predicates, and
 * whole-table deletion check live in `@stll/folio-core`; this hook owns only
 * the listener lifecycle, ref freshness, and the host-callback dispatch.
 */
export function useKeyboardShortcuts({
  pagedEditorRef,
  findReplace,
  tableSelection,
  onDirectPrint,
  scope,
  roots,
}: UseKeyboardShortcutsArgs): void {
  // Keep callbacks and roots fresh without re-attaching the listener on every
  // change to `findReplace.state` (which updates on every search keystroke) or
  // on a fresh `roots` array identity.
  const callbacksRef = useRef({ findReplace, tableSelection, onDirectPrint, roots });
  callbacksRef.current = { findReplace, tableSelection, onDirectPrint, roots };

  useEffect(() => {
    if (scope === "none") {
      return;
    }

    const handleDeleteSelectedTable = (e: KeyboardEvent) => {
      const view = pagedEditorRef.current?.getView();
      if (view && deleteSelectedTable(view.state, view.dispatch)) {
        e.preventDefault();
        return;
      }
      if (callbacksRef.current.tableSelection.state.tableIndex !== null) {
        e.preventDefault();
        callbacksRef.current.tableSelection.handleAction("deleteTable");
      }
    };

    const handleKeyDown = (e: KeyboardEvent) => {
      const inScope = isKeydownInShortcutScope(e, {
        scope,
        roots: callbacksRef.current.roots.map((root) => root.current),
      });
      if (!inScope) {
        return;
      }

      const editorDom = pagedEditorRef.current?.getView()?.dom;
      const intent = classifyEditorKeydown(e, {
        isMac: isMacPlatform(),
        isInputLike: isFocusInInputLike(e.target, editorDom),
      });

      switch (intent.type) {
        case "deleteSelectedTable":
          handleDeleteSelectedTable(e);
          return;
        case "openFind":
          e.preventDefault();
          callbacksRef.current.findReplace.openFind(readFindSelectionSeed());
          return;
        case "print":
          e.preventDefault();
          callbacksRef.current.onDirectPrint();
          return;
        case "none":
          return;
      }
    };

    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [pagedEditorRef, scope]);
}
