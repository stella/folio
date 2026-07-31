type ImeCaretHost = {
  style: {
    transform: string;
  };
};

type ImeCaretEditorView = {
  readonly composing: boolean;
  readonly dom?: HTMLElement;
  readonly state: {
    readonly selection: {
      readonly empty: boolean;
      readonly head: number;
    };
  };
  hasFocus: () => boolean;
  coordsAtPos: (position: number) => { left: number; top: number };
};

type VisibleCaretViewportRect = {
  left: number;
  top: number;
};

type SyncImeCaretAnchorOptions = {
  hiddenHost: ImeCaretHost | null | undefined;
  editorView: ImeCaretEditorView | null | undefined;
  visibleCaret: VisibleCaretViewportRect | null | undefined;
};

const resetHostTransform = (hiddenHost: ImeCaretHost | null | undefined): void => {
  if (hiddenHost) {
    hiddenHost.style.transform = "";
  }
};

const getNativeSelectionCaret = (
  editorView: ImeCaretEditorView,
): { left: number; top: number } | null => {
  const editorDom = editorView.dom;
  const selection = editorDom?.ownerDocument.getSelection();
  if (!editorDom || !selection || selection.rangeCount === 0) {
    return null;
  }

  const range = selection.getRangeAt(0);
  if (!editorDom.contains(range.startContainer)) {
    return null;
  }

  const rect = range.getBoundingClientRect();
  if (
    range.getClientRects().length === 0 ||
    (rect.left === 0 && rect.top === 0 && rect.width === 0 && rect.height === 0)
  ) {
    return null;
  }
  return { left: rect.left, top: rect.top };
};

/**
 * Align the real off-screen ProseMirror caret with the painted document caret.
 *
 * Native IME candidate windows follow the focused contenteditable selection,
 * not folio's visual caret overlay. Translating the hidden input host preserves
 * split rendering while anchoring CJK candidate UI at the visible insertion
 * point.
 */
export const syncImeCaretAnchor = ({
  hiddenHost,
  editorView,
  visibleCaret,
}: SyncImeCaretAnchorOptions): boolean => {
  if (!hiddenHost) {
    return false;
  }

  if (!editorView) {
    resetHostTransform(hiddenHost);
    return false;
  }

  // Moving a live composition target can dismiss or corrupt the native IME.
  // Keep the last valid anchor pinned until composition ends.
  if (editorView.composing) {
    return false;
  }

  if (!visibleCaret || !editorView.hasFocus() || !editorView.state.selection.empty) {
    resetHostTransform(hiddenHost);
    return false;
  }

  const previousTransform = hiddenHost.style.transform;
  hiddenHost.style.transform = "";

  let hiddenCaret: { left: number; top: number };
  try {
    // DOM selection geometry is the browser's actual native-UI anchor. Fall
    // back to ProseMirror geometry when the selection is temporarily stale.
    hiddenCaret =
      getNativeSelectionCaret(editorView) ??
      editorView.coordsAtPos(editorView.state.selection.head);
  } catch {
    hiddenHost.style.transform = previousTransform;
    return false;
  }

  if (
    !Number.isFinite(hiddenCaret.left) ||
    !Number.isFinite(hiddenCaret.top) ||
    !Number.isFinite(visibleCaret.left) ||
    !Number.isFinite(visibleCaret.top)
  ) {
    resetHostTransform(hiddenHost);
    return false;
  }

  const dx = Math.round(visibleCaret.left - hiddenCaret.left);
  const dy = Math.round(visibleCaret.top - hiddenCaret.top);
  hiddenHost.style.transform = `translate3d(${dx}px, ${dy}px, 0)`;
  return true;
};

export const resetImeCaretAnchor = (hiddenHost: ImeCaretHost | null | undefined): void => {
  resetHostTransform(hiddenHost);
};
