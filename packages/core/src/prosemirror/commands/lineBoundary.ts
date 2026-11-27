/**
 * Home / End as the start and end of the rendered line, on every platform.
 *
 * A browser's own Home and End differ by platform (on some, Home scrolls the
 * page and leaves the caret where it was), and a caret the browser moves is
 * read back only on a later `selectionchange`, after which a write of the
 * editor's selection can still put it back. This asks the browser for the
 * line boundary of the view's rendered lines and sets that position in the
 * editor's state as the key is handled.
 *
 * @packageDocumentation
 */
import { Selection, TextSelection } from "prosemirror-state";
import type { Command } from "prosemirror-state";

/** The DOM selection API the command needs, as browsers implement it. */
type LineSelection = Pick<
  globalThis.Selection,
  "focusNode" | "focusOffset" | "modify"
>;

/**
 * Move the caret, or with `extend` the selection's head, to the start
 * (`backward`) or end (`forward`) of the rendered line it is on.
 */
export const moveToLineBoundary =
  (direction: "backward" | "forward", extend: boolean): Command =>
  (state, dispatch, view) => {
    if (view === undefined) return false;
    const selection: LineSelection | null =
      view.dom.ownerDocument.getSelection();
    if (selection === null || typeof selection.modify !== "function")
      return false;
    const { focusNode } = selection;
    if (focusNode === null || !view.dom.contains(focusNode)) return false;
    selection.modify(extend ? "extend" : "move", direction, "lineboundary");
    const moved = selection.focusNode;
    if (moved === null || !view.dom.contains(moved)) return false;
    let head: number;
    try {
      head = view.posAtDOM(moved, selection.focusOffset);
    } catch {
      return false;
    }
    if (dispatch === undefined) return true;
    const { doc } = state;
    const anchor = extend ? state.selection.anchor : head;
    let next: Selection;
    try {
      next = TextSelection.create(doc, anchor, head);
    } catch {
      next = extend
        ? TextSelection.between(doc.resolve(anchor), doc.resolve(head))
        : Selection.near(doc.resolve(head));
    }
    dispatch(state.tr.setSelection(next).scrollIntoView());
    return true;
  };
