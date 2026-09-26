/**
 * How the user changes the document: `editing` writes directly, `suggesting`
 * records tracked changes, `viewing` is read-only.
 *
 * The same set as folio-react's `EditorMode`, spelled out so the protocol
 * module imports nothing from the editor and a host can bundle it alone.
 * `mount.tsx` passes modes both ways between the two types, so the compiler
 * fails if they drift apart.
 */
export type FolioEditingMode = "editing" | "suggesting" | "viewing";

const EDITING_MODES = {
  editing: true,
  suggesting: true,
  viewing: true,
} as const satisfies Record<FolioEditingMode, true>;

export const isEditingMode = (value: unknown): value is FolioEditingMode =>
  typeof value === "string" && Object.hasOwn(EDITING_MODES, value);
