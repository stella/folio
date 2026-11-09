import { scrollClipBoxToOrigin } from "../../paged-layout/editorScrollRoot";

/**
 * The class of the box hosts put the hidden editor's host in: a tiny clipping
 * box off-screen, which the browser may scroll to reveal the native caret.
 */
export const HIDDEN_EDITOR_WRAPPER_CLASS = "paged-editor__hidden-pm-wrapper";

type ImeCaretHost = {
  style: {
    transform: string;
  };
  readonly parentElement?: {
    closest: (selectors: string) => unknown;
  } | null;
};

/** The scroll state of the wrapper, as the browser keeps it. */
type ScrollingWrapper = {
  scrollLeft: number;
  scrollTop: number;
  addEventListener: (
    type: "scroll",
    listener: () => void,
    options?: { passive?: boolean },
  ) => void;
};

const isScrollingWrapper = (value: unknown): value is ScrollingWrapper =>
  typeof value === "object" &&
  value !== null &&
  "scrollLeft" in value &&
  "scrollTop" in value &&
  "addEventListener" in value;

const toOrigin = (wrapper: ScrollingWrapper): void =>
  scrollClipBoxToOrigin(wrapper);

const heldWrappers = new WeakSet<ScrollingWrapper>();

/**
 * Keep the wrapper around the host scrolled to its origin. The translation
 * moves the native caret outside the clipping wrapper, so the browser scrolls
 * the wrapper to reveal it, by any amount, whenever it reveals the caret; the
 * translation is measured with the wrapper at its origin and is only right
 * there. So the wrapper goes back to its origin before every measurement, and
 * after every scroll of it, which leaves the caret where the last measurement
 * put it.
 */
const holdWrapperAtOrigin = (hiddenHost: ImeCaretHost): void => {
  const wrapper = hiddenHost.parentElement?.closest(
    `.${HIDDEN_EDITOR_WRAPPER_CLASS}`,
  );
  if (!isScrollingWrapper(wrapper)) return;
  toOrigin(wrapper);
  if (heldWrappers.has(wrapper)) return;
  heldWrappers.add(wrapper);
  wrapper.addEventListener("scroll", () => toOrigin(wrapper), {
    passive: true,
  });
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

/** A translation this module wrote, as the browser or a test reads it back. */
const ANCHOR_TRANSLATION =
  /^translate3d\((-?\d+(?:\.\d+)?)px, (-?\d+(?:\.\d+)?)px, 0(?:px)?\)$/u;

type HostOffset = { dx: number; dy: number };

/**
 * The offset the host carries now: none without a transform, the translation
 * for one this module wrote, null for any other transform.
 */
const hostOffset = (hiddenHost: ImeCaretHost): HostOffset | null => {
  const transform = hiddenHost.style.transform;
  if (transform === "") return { dx: 0, dy: 0 };
  const match = ANCHOR_TRANSLATION.exec(transform);
  if (!match) return null;
  return { dx: Number(match[1]), dy: Number(match[2]) };
};

/**
 * Give the host an offset, writing only when it changes. A reset keeps a zero
 * translation rather than removing the transform: the host holds the whole
 * hidden document, and adding or removing its transform makes the browser
 * rebuild paint properties for every node under it.
 */
const placeHost = (hiddenHost: ImeCaretHost, dx: number, dy: number): void => {
  const current = hostOffset(hiddenHost);
  if (current && current.dx === dx && current.dy === dy) return;
  hiddenHost.style.transform = `translate3d(${dx}px, ${dy}px, 0)`;
};

const resetHostTransform = (
  hiddenHost: ImeCaretHost | null | undefined,
): void => {
  if (hiddenHost) {
    placeHost(hiddenHost, 0, 0);
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
  holdWrapperAtOrigin(hiddenHost);

  if (!editorView) {
    resetHostTransform(hiddenHost);
    return false;
  }

  // Moving a live composition target can dismiss or corrupt the native IME.
  // Keep the last valid anchor pinned until composition ends.
  if (editorView.composing) {
    return false;
  }

  if (
    !visibleCaret ||
    !editorView.hasFocus() ||
    !editorView.state.selection.empty
  ) {
    resetHostTransform(hiddenHost);
    return false;
  }

  // Measure the hidden caret where it stands and take off the offset the host
  // carries, instead of clearing the transform to measure and setting it again.
  const previousTransform = hiddenHost.style.transform;
  let offset = hostOffset(hiddenHost);
  if (!offset) {
    hiddenHost.style.transform = "";
    offset = { dx: 0, dy: 0 };
  }

  let hiddenCaret: { left: number; top: number };
  try {
    // DOM selection geometry is the browser's actual native-UI anchor. Fall
    // back to ProseMirror geometry when the selection is temporarily stale.
    hiddenCaret =
      getNativeSelectionCaret(editorView) ??
      editorView.coordsAtPos(editorView.state.selection.head);
  } catch {
    if (hiddenHost.style.transform !== previousTransform)
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

  const dx = Math.round(visibleCaret.left - (hiddenCaret.left - offset.dx));
  const dy = Math.round(visibleCaret.top - (hiddenCaret.top - offset.dy));
  placeHost(hiddenHost, dx, dy);
  return true;
};

export const resetImeCaretAnchor = (
  hiddenHost: ImeCaretHost | null | undefined,
): void => {
  resetHostTransform(hiddenHost);
};
