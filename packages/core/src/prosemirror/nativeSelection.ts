import { Plugin, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

/** Take the caret the browser shows into the editor state before input. */
export const adoptNativeSelection = (view: EditorView): void => {
  // Keys the browser handles itself (End, Home, the arrows, the page keys) move
  // the DOM selection, and ProseMirror reads that move from the asynchronous
  // `selectionchange` event, which can arrive after the next key. Input the
  // editor applies at its own selection (typed text, Enter, Backspace) would
  // then land where the caret was before the move. Every key and text input
  // takes the caret the browser shows first.
  //
  // A text selection follows the DOM. Any other selection (select-all, a node,
  // table cells) follows it only to a collapsed caret in text, which a
  // navigation key leaves and which cannot show that selection; the range the
  // browser paints for it keeps it. ProseMirror reads the end of one run and the
  // start of the next as one position, so a caret on either side of a mark
  // boundary leaves the selection, and its stored marks, as they are.
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

const nativeRange = (view: EditorView): Range | null => {
  const native = view.dom.ownerDocument.getSelection();
  return native !== null && native.rangeCount > 0 ? native.getRangeAt(0) : null;
};

/** The image ProseMirror draws beside an inline atom with no text there. */
const isElement = (node: Node): node is Element =>
  node.nodeType === node.ELEMENT_NODE;
const isSeparator = (node: Node | undefined): boolean =>
  node !== undefined &&
  isElement(node) &&
  node.nodeName === "IMG" &&
  node.classList.contains("ProseMirror-separator");

/**
 * Where a browser caret for the editor's collapsed selection belongs: in the
 * text before it, or at a paragraph's start in the first text. ProseMirror
 * treats the end of one text node and the start of the next run as one
 * position, but the browser types and composes into the node it holds.
 *
 * With no text beside it, the caret sits before or after an inline atom,
 * such as a bookmark boundary at a paragraph's start, and ProseMirror draws
 * a separator image there. The caret goes right after that image: before it
 * and the atom, Chromium finds no place to compose and composes into the
 * first text of the editable content instead.
 *
 * `null` when the selection is not a caret in text or beside such an image.
 */
export const textCaretOf = (
  view: EditorView,
): { node: Node; offset: number } | null => {
  const { selection } = view.state;
  if (!selection.empty) return null;
  const sides = [
    view.domAtPos(selection.from, -1),
    view.domAtPos(selection.from, 1),
  ];
  const text = sides.find(({ node }) => node.nodeType === node.TEXT_NODE);
  if (text) return { node: text.node, offset: text.offset };
  for (const { node, offset } of sides) {
    if (isSeparator(node.childNodes[offset - 1])) return { node, offset };
    if (isSeparator(node.childNodes[offset]))
      return { node, offset: offset + 1 };
  }
  return null;
};

/**
 * Sets the browser's selection again after an update that changed the
 * document, unless it was just set where it belongs.
 *
 * ProseMirror leaves a DOM selection it reads as the same position, such as
 * the start of the next run instead of the end of the previous one, or a
 * caret beside a redrawn paragraph. Chromium's input method then resolves the
 * next composition's start from the start of the editable content, and it
 * picks the text node to compose into before `compositionstart`, so the
 * text in progress shows elsewhere. A caret goes into the text before it,
 * where a composition continues the run it follows.
 *
 * A selection the browser or ProseMirror set since the last update has a new
 * range; one left alone keeps its range and is set again.
 */
export const createNativeSelectionRefreshPlugin = (): Plugin =>
  new Plugin({
    view: (view) => {
      let held = nativeRange(view);
      return {
        update: (current, previous) => {
          const native = current.dom.ownerDocument.getSelection();
          const range = nativeRange(current);
          if (
            native !== null &&
            range !== null &&
            current.state.doc !== previous.doc &&
            !current.composing &&
            current.hasFocus() &&
            current.dom.contains(range.startContainer) &&
            current.dom.contains(range.endContainer)
          ) {
            const stale = range === held;
            const caret = native.isCollapsed ? textCaretOf(current) : null;
            const { anchorNode, anchorOffset, focusNode, focusOffset } = native;
            if (
              caret !== null &&
              (stale ||
                anchorNode !== caret.node ||
                anchorOffset !== caret.offset)
            ) {
              native.removeAllRanges();
              native.collapse(caret.node, caret.offset);
            } else if (caret === null && stale && anchorNode && focusNode) {
              // The same selection set again counts only after a change.
              native.removeAllRanges();
              native.setBaseAndExtent(
                anchorNode,
                anchorOffset,
                focusNode,
                focusOffset,
              );
            }
          }
          held = nativeRange(current);
        },
      };
    },
  });
