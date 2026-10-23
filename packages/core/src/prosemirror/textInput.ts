import type { Node as ProseMirrorNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import {
  AllSelection,
  Plugin,
  Selection,
  TextSelection,
} from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";

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

type FormattingMarksBesideLockedControlOptions = {
  state: EditorState;
  from: number;
  to: number;
};

const LOCKED_CONTROL_LOCKS = new Set(["contentLocked", "sdtContentLocked"]);

const isLockedControl = (node: ProseMirrorNode | null) =>
  node?.type.name === "sdt" && LOCKED_CONTROL_LOCKS.has(node.attrs["lock"]);

export const formattingMarksBesideLockedControl = ({
  state,
  from,
  to,
}: FormattingMarksBesideLockedControlOptions) => {
  if (from !== to || state.storedMarks !== null) return undefined;
  const $from = state.doc.resolve(from);
  const { parent } = $from;
  if (!parent.isTextblock) return undefined;
  const before = $from.nodeBefore;
  const after = $from.nodeAfter;
  const control = [before, after].find(isLockedControl);
  if (!control) return undefined;
  const afterControl = before === control;
  const controlIndex = afterControl ? $from.index() - 1 : $from.index();
  const adjacent = parent.maybeChild(controlIndex + (afterControl ? 1 : -1));
  if (
    !adjacent?.isText ||
    adjacent.marks.some(
      (mark) => mark.type.name === "deletion" || mark.type.name === "moveFrom",
    )
  ) {
    return [];
  }
  return adjacent.marks.filter((mark) =>
    RUN_FORMATTING_MARK_NAMES.has(mark.type.name),
  );
};

export const dispatchEditorTextInput = <
  TView extends TextInputDispatchTarget<TView>,
>(
  view: TView,
  input: string | TextInput,
) => {
  const { from, to } = typeof input === "string" ? view.state.selection : input;
  const text = typeof input === "string" ? input : input.text;
  const defaultTransaction = () => {
    const { selection } = view.state;
    if (!(selection instanceof AllSelection)) {
      const formattingMarks = formattingMarksBesideLockedControl({
        state: view.state,
        from,
        to,
      });
      const tr = view.state.tr;
      if (formattingMarks !== undefined) {
        tr.setStoredMarks(formattingMarks);
      }
      tr.insertText(text, from, to);
      // Typing over a selection leaves the caret after the typed text;
      // `insertText` puts it at the selection's mapped end, which for a node
      // selection (an image, a field) is before the text.
      return !selection.empty && selection.from === from && selection.to === to
        ? tr.setSelection(TextSelection.create(tr.doc, from + text.length))
        : tr;
    }
    // Select-all spans the block boundaries around the content: typing
    // replaces the text inside them, formatted as the first character it
    // replaces, and leaves the caret after it. The paragraph left is the last
    // one, whose break is the one that stays.
    const start = Selection.atStart(view.state.doc);
    const end = Selection.atEnd(view.state.doc);
    const last = end.$to.parent;
    const tr = view.state.tr.insertText(text, start.from, end.to);
    const $typed = tr.doc.resolve(start.from);
    if (
      start.$from.parent.type.name === "paragraph" &&
      last.type.name === "paragraph" &&
      $typed.parent.type.name === "paragraph" &&
      start.$from.parent !== last
    ) {
      tr.setNodeMarkup($typed.before(), undefined, last.attrs);
    }
    return tr.setSelection(
      TextSelection.create(tr.doc, start.from + text.length),
    );
  };
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
export const getEditorInputTargetRange = (
  view: EditorView,
  event: InputEvent,
) => {
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
export const handleEditorBeforeInput = (
  view: EditorView,
  event: InputEvent,
): boolean => {
  if (
    !view.editable ||
    !event.cancelable ||
    event.defaultPrevented ||
    view.composing ||
    event.isComposing
  ) {
    return false;
  }
  if (
    event.inputType !== "insertText" &&
    event.inputType !== "insertReplacementText"
  ) {
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
      ? getEditorInputTargetRange(view, event)
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

/**
 * Plugin compositionend handlers run before ProseMirror queues its final DOM
 * flush. A second microtask runs after that flush and sees the committed state.
 */
export const afterNativeCompositionFlush = (callback: () => void): void => {
  queueMicrotask(() => queueMicrotask(callback));
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
          afterNativeCompositionFlush(() => {
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
