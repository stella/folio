import { useCallback, useMemo, useSyncExternalStore } from "react";
import type { RefObject } from "react";

import type { EditorState } from "prosemirror-state";

import { ContextMenuManager } from "@stll/folio-core/managers/ContextMenuManager";
import type {
  ContextMenuAnchor,
  ContextMenuSnapshot,
  ContextMenuTrackedChange,
} from "@stll/folio-core/managers/types";

import type { PagedEditorRef } from "../../paged-editor/PagedEditor";
import { detectTrackedChangeAt } from "../selectionDetection";

export type { ContextMenuAnchor } from "@stll/folio-core/managers/types";
export type ContextMenuState = ContextMenuSnapshot;

export type UseContextMenuArgs = {
  pagedEditorRef: RefObject<PagedEditorRef | null>;
};

/** The story state and position under the pointer that opened the menu. */
export type ContextMenuPointer = { state: EditorState; position: number };

type OpenMenuOptions = {
  hasSelectionOverride?: boolean | undefined;
  /** Null when the pointer is on no position of an editable story. */
  pointer: ContextMenuPointer | null;
};

/**
 * The tracked change a menu opened at `pointer` resolves: the one under the
 * pointer, never the one at the selection, since a right-click inside a range
 * selection keeps the range wherever it starts.
 */
export const contextMenuTrackedChangeAt = (
  pointer: ContextMenuPointer | null,
): ContextMenuTrackedChange | null => {
  const change =
    pointer === null
      ? null
      : detectTrackedChangeAt(pointer.state, pointer.position);
  return (
    change && {
      from: change.from,
      to: change.to,
      heldByLock: change.heldByLock,
    }
  );
};

export type UseContextMenuReturn = {
  contextMenu: ContextMenuState;
  /**
   * Open the menu at `anchor`. Reads the live PM selection from
   * `pagedEditorRef` to decide whether to enable selection-only and
   * table-only items, which act on the selection; the tracked-change items
   * act on the change under `pointer`.
   *
   * `hasSelectionOverride` is for the PagedEditor child callback path,
   * which already knows from a layout-overlay selection whether the user
   * has highlighted text; when omitted we derive it from the PM selection.
   */
  openMenu: (anchor: ContextMenuAnchor, options: OpenMenuOptions) => void;
  closeMenu: () => void;
};

export function useContextMenu({
  pagedEditorRef,
}: UseContextMenuArgs): UseContextMenuReturn {
  const manager = useMemo(() => new ContextMenuManager(), []);
  const contextMenu = useSyncExternalStore(
    manager.subscribe,
    manager.getSnapshot,
  );

  const openMenu = useCallback(
    (anchor: ContextMenuAnchor, options: OpenMenuOptions) => {
      const state = pagedEditorRef.current?.getView()?.state ?? null;
      manager.openMenu(anchor, {
        state,
        hasSelectionOverride: options.hasSelectionOverride,
        trackedChange: contextMenuTrackedChangeAt(options.pointer),
      });
    },
    [pagedEditorRef, manager],
  );

  const closeMenu = useCallback(() => {
    manager.closeMenu();
  }, [manager]);

  return { contextMenu, openMenu, closeMenu };
}
