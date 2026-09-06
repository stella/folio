/**
 * The search text a find/replace dialog opens with: the current DOM selection,
 * or `""` when it is collapsed or absent. One helper so the keyboard shortcut
 * and `DocxEditorRef.openFind` / `openReplace` cannot seed differently.
 */
export const readFindSelectionSeed = (): string => {
  const selection = window.getSelection();
  return selection && !selection.isCollapsed ? selection.toString() : "";
};
