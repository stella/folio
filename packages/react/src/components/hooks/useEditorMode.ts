import { useCallback, useState, useSyncExternalStore } from "react";

import {
  EditorModeManager,
  isReadOnlyMode,
  isTrackChangesMode,
  toggledTrackChangesMode,
} from "@stll/folio-core/managers/EditorModeManager";
import type { DisplayMode, EditorMode } from "@stll/folio-core/managers/EditorModeManager";

import { reviewDisplayModeToDisplayMode } from "@stll/folio-core/managers/reviewDisplayMode";
import type { ReviewDisplayMode } from "@stll/folio-core/managers/reviewDisplayMode";

export { DISPLAY_MODES } from "@stll/folio-core/managers/EditorModeManager";
export type { DisplayMode, EditorMode } from "@stll/folio-core/managers/EditorModeManager";

export type UseEditorModeArgs = {
  /** Controlled mode prop, or undefined when uncontrolled. */
  modeProp: EditorMode | undefined;
  /** Notified whenever the mode changes (controlled and uncontrolled paths). */
  onModeChange: ((mode: EditorMode) => void) | undefined;
  /** External read-only flag from the host. */
  readOnlyProp: boolean;
  reviewDisplayMode?: ReviewDisplayMode;
};

export type UseEditorModeReturn = {
  /** Effective editing mode (controlled `modeProp` wins over the internal value). */
  editingMode: EditorMode;
  /** Change mode. No-op for the internal state when the host controls it via `modeProp`. */
  setEditingMode: (mode: EditorMode) => void;
  /** True when the host opts into read-only OR the mode is `"viewing"`. */
  readOnly: boolean;
  /** True when the mode is `"suggesting"`. */
  trackChangesOn: boolean;
  /** Toggle between `"editing"` and `"suggesting"`. */
  toggleTrackChanges: () => void;
  /** Display mode for the tracked-changes overlay (`all-markup` by default). */
  displayMode: DisplayMode;
  setDisplayMode: (mode: DisplayMode) => void;
};

/**
 * Editor-mode and display-mode state for the DocxEditor.
 *
 * Thin React binding around the framework-agnostic EditorModeManager. The
 * manager owns the uncontrolled internal state; this hook layers the
 * controlled-prop reconciliation (`modeProp` wins) and the `onModeChange`
 * notification on top.
 *
 * `editingMode` mirrors Google Docs:
 *  - `"editing"`     → direct edits (default)
 *  - `"suggesting"`  → tracked-change edits
 *  - `"viewing"`     → behaves as read-only
 *
 * `displayMode` is independent: it controls how tracked-changes render
 * (`all-markup`, `simple-markup`, `no-markup`, `original`). Switching the
 * editing mode does not switch the display mode.
 */
export function useEditorMode({
  modeProp,
  onModeChange,
  readOnlyProp,
  reviewDisplayMode,
}: UseEditorModeArgs): UseEditorModeReturn {
  const [manager] = useState(() => new EditorModeManager(modeProp ?? "editing"));
  const { editingMode: editingModeInternal, displayMode } = useSyncExternalStore(
    manager.subscribe,
    manager.getSnapshot,
  );
  const editingMode = modeProp ?? editingModeInternal;

  const setEditingMode = useCallback(
    (mode: EditorMode) => {
      if (!modeProp) {
        manager.setEditingMode(mode);
      }
      onModeChange?.(mode);
    },
    [modeProp, onModeChange, manager],
  );

  const readOnly = isReadOnlyMode(editingMode, readOnlyProp);
  const trackChangesOn = isTrackChangesMode(editingMode);

  const toggleTrackChanges = useCallback(() => {
    setEditingMode(toggledTrackChangesMode(editingMode));
  }, [setEditingMode, editingMode]);

  const setDisplayMode = useCallback(
    (mode: DisplayMode) => {
      manager.setDisplayMode(mode);
    },
    [manager],
  );

  return {
    editingMode,
    setEditingMode,
    readOnly,
    trackChangesOn,
    toggleTrackChanges,
    displayMode:
      reviewDisplayMode === undefined
        ? displayMode
        : reviewDisplayModeToDisplayMode(reviewDisplayMode),
    setDisplayMode,
  };
}
