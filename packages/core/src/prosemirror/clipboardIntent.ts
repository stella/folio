import type { Slice } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";

/** Async clipboard reads resolve against the current owning session and selection. */
const handlers = new WeakMap<EditorView, (slice: Slice, plain: boolean) => boolean>();

export const registerClipboardIntentHandler = (
  view: EditorView,
  handler: (slice: Slice, plain: boolean) => boolean,
): void => {
  handlers.set(view, handler);
};

export const dispatchClipboardIntent = (view: EditorView, slice: Slice): boolean | undefined =>
  handlers.get(view)?.(slice, true);
