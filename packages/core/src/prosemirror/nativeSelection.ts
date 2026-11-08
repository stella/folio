import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

/**
 * Keys the browser handles itself (End, Home, the arrows, the page keys) move
 * the DOM selection, and ProseMirror reads that move from the asynchronous
 * `selectionchange` event, which can arrive after the next key. Input the
 * editor applies at its own selection (typed text, Enter, Backspace) would
 * then land where the caret was before the move. Every key and text input
 * takes the caret the browser shows first.
 *
 * A text selection follows the DOM. Any other selection (select-all, a node,
 * table cells) follows it only to a collapsed caret in text, which a
 * navigation key leaves and which cannot show that selection; the range the
 * browser paints for it keeps it. ProseMirror reads the end of one run and the
 * start of the next as one position, so a caret on either side of a mark
 * boundary leaves the selection, and its stored marks, as they are.
 */
export const adoptNativeSelection = (view: EditorView): void => {
  if (view.composing || view.isDestroyed) return;
  const { selection } = view.state;
  const native = view.dom.ownerDocument.getSelection();
  const follows =
    selection instanceof TextSelection || native?.isCollapsed === true;
  if (
    !follows ||
    !native?.anchorNode ||
    !native.focusNode ||
    !view.dom.contains(native.anchorNode) ||
    !view.dom.contains(native.focusNode)
  )
    return;
  const anchor = view.posAtDOM(native.anchorNode, native.anchorOffset);
  const head = view.posAtDOM(native.focusNode, native.focusOffset);
  if (anchor === selection.anchor && head === selection.head) return;
  const { doc } = view.state;
  if (
    !doc.resolve(anchor).parent.inlineContent ||
    !doc.resolve(head).parent.inlineContent
  )
    return;
  view.dispatch(
    view.state.tr.setSelection(TextSelection.create(doc, anchor, head)),
  );
};

/** Whether a keydown belongs to an input method composition. */
export const isCompositionKeyDown = (event: KeyboardEvent): boolean =>
  event.isComposing || event.key === "Process";
