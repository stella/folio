import { useCallback, useMemo, useSyncExternalStore } from "react";
import type { RefObject } from "react";

import { ContextMenuManager } from "@stll/folio-core/managers/ContextMenuManager";
import type {
  ContextMenuAnchor,
  ContextMenuSnapshot,
} from "@stll/folio-core/managers/types";

import type { PagedEditorRef } from "../../paged-editor/PagedEditor";
import { detectActiveTrackedChange } from "../selectionDetection";

export type { ContextMenuAnchor } from "@stll/folio-core/managers/types";
export type ContextMenuState = ContextMenuSnapshot;

export type UseContextMenuArgs = {
  pagedEditorRef: RefObject<PagedEditorRef | null>;
};

export type UseContextMenuReturn = {
  contextMenu: ContextMenuState;
  /**
   * Open the menu at `anchor`. Reads the live PM selection from
   * `pagedEditorRef` to decide whether to enable selection-only,
   * table-only, and tracked-change-only menu items.
   *
   * `hasSelectionOverride` is for the PagedEditor child callback path,
   * which already knows from a layout-overlay selection whether the user
   * has highlighted text; when omitted we derive it from the PM selection.
   */
  openMenu: (anchor: ContextMenuAnchor, hasSelectionOverride?: boolean) => void;
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
    (anchor: ContextMenuAnchor, hasSelectionOverride?: boolean) => {
      const state = pagedEditorRef.current?.getView()?.state ?? null;
      const cursorInTrackedChange = state
        ? detectActiveTrackedChange(state) !== null
        : false;
      manager.openMenu(anchor, {
        state,
        hasSelectionOverride,
        cursorInTrackedChange,
      });
    },
    [pagedEditorRef, manager],
  );

  const closeMenu = useCallback(() => {
    manager.closeMenu();
  }, [manager]);

  return { contextMenu, openMenu, closeMenu };
}
