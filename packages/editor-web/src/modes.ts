import type { EditorMode } from "@stll/folio-react";

/**
 * How the user changes the document: `editing` writes directly, `suggesting`
 * records tracked changes, `viewing` is read-only.
 */
export type FolioEditingMode = EditorMode;

const EDITING_MODES = {
  editing: true,
  suggesting: true,
  viewing: true,
} as const satisfies Record<FolioEditingMode, true>;

export const isEditingMode = (value: unknown): value is FolioEditingMode =>
  typeof value === "string" && Object.hasOwn(EDITING_MODES, value);
