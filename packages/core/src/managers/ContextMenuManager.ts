/**
 * ContextMenuManager
 *
 * Framework-agnostic store for the editor context-menu state.
 * Extracted from the React `useContextMenu` hook.
 *
 * `openMenu` derives the selection-, table-, and tracked-change-aware menu
 * flags from the live ProseMirror state. The tracked-change flag is passed in
 * by the adapter because its detection currently lives in the React layer; the
 * rest is computed here from core's ProseMirror queries.
 */

import type { EditorState } from "prosemirror-state";

import { isInTable } from "../prosemirror/commands/table";
import { listNumberingMenuState, NO_LIST_NUMBERING_MENU } from "../prosemirror/listNumbering";
import { Subscribable } from "./Subscribable";
import type { ContextMenuAnchor, ContextMenuSnapshot } from "./types";

const CLOSED_STATE: ContextMenuSnapshot = {
  isOpen: false,
  position: { x: 0, y: 0 },
  hasSelection: false,
  selectionRange: { from: 0, to: 0 },
  cursorInTable: false,
  cursorInTrackedChange: false,
  listNumbering: NO_LIST_NUMBERING_MENU,
};

export type OpenContextMenuOptions = {
  /** Live editor state, or null when the editor view is not mounted. */
  state: EditorState | null;
  /**
   * For the PagedEditor child callback path, which already knows from a
   * layout-overlay selection whether the user highlighted text. When omitted,
   * `hasSelection` is derived from the ProseMirror selection.
   */
  hasSelectionOverride?: boolean | undefined;
  /** Whether the cursor sits inside a tracked change (detected by the adapter). */
  cursorInTrackedChange: boolean;
};

export class ContextMenuManager extends Subscribable<ContextMenuSnapshot> {
  constructor() {
    super(CLOSED_STATE);
  }

  /** Open the menu at `anchor`, deriving menu flags from the editor state. */
  openMenu(
    anchor: ContextMenuAnchor,
    { state, hasSelectionOverride, cursorInTrackedChange }: OpenContextMenuOptions,
  ): void {
    const selection = state?.selection ?? { from: 0, to: 0 };
    const hasSelection = hasSelectionOverride ?? selection.from !== selection.to;
    const cursorInTable = state ? isInTable(state) : false;

    this.setSnapshot({
      isOpen: true,
      position: anchor,
      hasSelection,
      selectionRange: { from: selection.from, to: selection.to },
      cursorInTable,
      cursorInTrackedChange,
      listNumbering: state ? listNumberingMenuState(state) : NO_LIST_NUMBERING_MENU,
    });
  }

  /** Close the menu. */
  closeMenu(): void {
    this.setSnapshot(CLOSED_STATE);
  }
}
