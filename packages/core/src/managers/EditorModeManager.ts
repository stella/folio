/**
 * EditorModeManager
 *
 * Framework-agnostic store for the editor's mode state.
 * Extracted from the React `useEditorMode` hook.
 *
 * Holds the uncontrolled internal editing mode and the tracked-changes display
 * mode; the host adapter layers controlled-prop reconciliation on top. The
 * mode derivations (read-only, track-changes, toggle target) are pure helpers
 * so any adapter can compute them without the store.
 */

import { Subscribable } from "./Subscribable";

// ============================================================================
// CONSTANTS & TYPES
// ============================================================================

/** Editor mode. Mirrors Google Docs editing/suggesting/viewing semantics. */
export type EditorMode = "editing" | "suggesting" | "viewing";

/** How tracked changes render. Drives the display mode dropdown. */
export const DISPLAY_MODES = ["all-markup", "simple-markup", "no-markup", "original"] as const;

export type DisplayMode = (typeof DISPLAY_MODES)[number];

/** Why a user action cannot edit the displayed document. */
export type EditorReadOnlyReason = "host-read-only" | "viewing-mode" | "review-display-mode";

const REVIEW_VIEW_ACCESS = {
  "all-markup": "editable",
  "simple-markup": "read-only",
  "no-markup": "read-only",
  original: "read-only",
} as const satisfies Record<DisplayMode, "editable" | "read-only">;

export const getReviewViewReadOnlyReason = (
  displayMode: DisplayMode,
): EditorReadOnlyReason | null =>
  REVIEW_VIEW_ACCESS[displayMode] === "read-only" ? "review-display-mode" : null;

type EditorReadOnlyReasonOptions = {
  editingMode: EditorMode;
  readOnlyProp: boolean;
  displayMode: DisplayMode;
};

export const getEditorReadOnlyReason = ({
  editingMode,
  readOnlyProp,
  displayMode,
}: EditorReadOnlyReasonOptions): EditorReadOnlyReason | null => {
  const viewReason = getReviewViewReadOnlyReason(displayMode);
  if (viewReason) return viewReason;
  if (readOnlyProp) return "host-read-only";
  if (editingMode === "viewing") return "viewing-mode";
  return null;
};

/** EditorModeManager snapshot. */
export type EditorModeSnapshot = {
  /** Uncontrolled internal editing mode (the host's controlled prop wins over this). */
  editingMode: EditorMode;
  /** Tracked-changes display mode (`all-markup` by default). */
  displayMode: DisplayMode;
};

// ============================================================================
// PURE DERIVATIONS
// ============================================================================

/** True when the host opts into read-only OR the mode is `"viewing"`. */
export const isReadOnlyMode = (editingMode: EditorMode, readOnlyProp: boolean): boolean =>
  readOnlyProp || editingMode === "viewing";

/** True when the mode is `"suggesting"`. */
export const isTrackChangesMode = (editingMode: EditorMode): boolean =>
  editingMode === "suggesting";

/** The mode produced by toggling track changes between `"editing"` and `"suggesting"`. */
export const toggledTrackChangesMode = (editingMode: EditorMode): EditorMode =>
  isTrackChangesMode(editingMode) ? "editing" : "suggesting";

// ============================================================================
// MANAGER
// ============================================================================

export class EditorModeManager extends Subscribable<EditorModeSnapshot> {
  constructor(initialEditingMode: EditorMode) {
    super({ editingMode: initialEditingMode, displayMode: "all-markup" });
  }

  /** Set the internal editing mode. */
  setEditingMode(mode: EditorMode): void {
    this.setSnapshot({ ...this.getSnapshot(), editingMode: mode });
  }

  /** Set the tracked-changes display mode. */
  setDisplayMode(mode: DisplayMode): void {
    this.setSnapshot({ ...this.getSnapshot(), displayMode: mode });
  }
}
