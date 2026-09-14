import type { EditorState, Transaction } from "prosemirror-state";
import { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

type TextInputHandler<TView> = (
  view: TView,
  from: number,
  to: number,
  text: string,
  defaultTransaction: () => Transaction,
) => unknown;

type TextInputDispatchTarget<TView> = {
  dispatch: (tr: Transaction) => void;
  someProp: (
    propName: "handleTextInput",
    f: (handler: TextInputHandler<TView>) => unknown,
  ) => unknown;
  state: EditorState;
};

type TextInput = { text: string; from: number; to: number };

export const dispatchEditorTextInput = <TView extends TextInputDispatchTarget<TView>>(
  view: TView,
  input: string | TextInput,
) => {
  const { from, to } = typeof input === "string" ? view.state.selection : input;
  const text = typeof input === "string" ? input : input.text;
  const defaultTransaction = () => view.state.tr.insertText(text, from, to);
  const handled = view.someProp("handleTextInput", (handler) =>
    handler(view, from, to, text, defaultTransaction),
  );

  if (!handled) {
    view.dispatch(defaultTransaction());
  }
};

// Browser replacement operations (spellcheck/autocorrect) may target text other
// than the current selection. Ordinary typing uses the model selection, whose
// mark boundary is unambiguous even when the DOM caret is inside the left span.
const replacementSelection = (view: EditorView, event: InputEvent) => {
  if (typeof event.getTargetRanges !== "function") return null;
  const ranges = event.getTargetRanges();
  if (ranges.length !== 1) return null;
  const range = ranges[0];
  if (
    !range ||
    !view.dom.contains(range.startContainer) ||
    !view.dom.contains(range.endContainer)
  ) {
    return null;
  }
  const from = view.posAtDOM(range.startContainer, range.startOffset);
  const to = view.posAtDOM(range.endContainer, range.endOffset);
  if (from < 0 || to < from || to > view.state.doc.content.size) return null;
  return { from, to };
};

/** Shared DOM boundary for editor runtimes and standalone suggestion views. */
export const handleEditorBeforeInput = (view: EditorView, event: InputEvent): boolean => {
  if (
    !view.editable ||
    !event.cancelable ||
    event.defaultPrevented ||
    view.composing ||
    event.isComposing
  ) {
    return false;
  }
  if (event.inputType !== "insertText" && event.inputType !== "insertReplacementText") {
    return false;
  }
  const text =
    event.data ??
    (event.dataTransfer?.types.includes("text/plain")
      ? event.dataTransfer.getData("text/plain")
      : null);
  if (text === null) return false;
  const selection =
    event.inputType === "insertReplacementText"
      ? replacementSelection(view, event)
      : view.state.selection;
  if (!selection) return false;
  event.preventDefault();
  dispatchEditorTextInput(view, {
    text,
    from: selection.from,
    to: selection.to,
  });
  return true;
};

/** Route committed text through the model before the browser mutates mark spans. */
export const createTextInputPlugin = () => {
  // PM can clear view.composing before its final DOM flush. Keep that commit
  // native through the end-of-composition microtask, just like suggestion mode.
  // Key by view because a plugin can be shared by multiple editor states/views.
  const compositions = new WeakMap<EditorView, Event>();
  return new Plugin({
    props: {
      handleDOMEvents: {
        compositionstart(view, event) {
          compositions.set(view, event);
          return false;
        },
        compositionend(view, event) {
          compositions.set(view, event);
          queueMicrotask(() => {
            if (compositions.get(view) === event) compositions.delete(view);
          });
          return false;
        },
        beforeinput(view, event) {
          if (compositions.has(view)) return false;
          return handleEditorBeforeInput(view, event);
        },
      },
    },
  });
};
