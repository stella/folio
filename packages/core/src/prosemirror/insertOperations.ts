/**
 * Insert Operations
 *
 * View-level helpers that run the editor's structural insert commands against a
 * live `EditorView`. They are thin ergonomic wrappers over the registered
 * ProseMirror commands (`insertTable`, `insertPageBreak`, `generateTOC`) and the
 * image-from-file command, so a consumer wiring toolbar handlers does not have
 * to know each command's calling convention (factory vs. plain `Command`).
 */

import type { EditorView } from "prosemirror-view";

import { generateTOC, insertPageBreak, insertTable } from "./commands";
import { insertImageFromFile } from "./commands/image";

export { insertImageFromFile };

/** Insert a `rows × columns` table at the current selection. */
export function insertTableInView(view: EditorView, rows: number, columns: number): boolean {
  return insertTable(rows, columns)(view.state, view.dispatch);
}

/** Insert a page break at the current selection. */
export function insertPageBreakInView(view: EditorView): boolean {
  return insertPageBreak(view.state, view.dispatch);
}

/** Generate (or refresh) a table of contents from the document's headings. */
export function insertTableOfContentsInView(view: EditorView): boolean {
  return generateTOC(view.state, view.dispatch);
}
