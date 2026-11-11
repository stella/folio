import type { Mark, Node as ProseMirrorNode } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import {
  AllSelection,
  Plugin,
  Selection,
  TextSelection,
} from "prosemirror-state";
import { Mapping } from "prosemirror-transform";
import type { EditorView } from "prosemirror-view";

import { adoptNativeSelection, isCompositionKeyDown } from "./nativeSelection";
import { createPluginKeyOwner } from "./pluginKeyOwner";
import { RUN_FORMATTING_MARK_NAMES } from "./runFormattingMarkNames";
import { isZeroWidthAnchor } from "./zeroWidthAnchors";

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
}: FormattingMarksBesideLockedControlOptions): Mark[] | undefined => {
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
  const adjacent =
    parent.maybeChild(controlIndex + (afterControl ? 1 : -1)) ??
    parent.maybeChild(controlIndex + (afterControl ? -1 : 1));
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

const isCommentMark = (mark: Mark): boolean => mark.type.name === "comment";

/**
 * Whether `node` is content a comment can cover: text, or an inline leaf that
 * shows something (an image, a field, a tab). Anchors that only hold a
 * position (comment references, range and bookmark boundaries) are not.
 */
const isCoverableContent = (node: ProseMirrorNode): boolean =>
  node.isText ||
  (node.isInline && (node.isLeaf || node.isAtom) && !isZeroWidthAnchor(node));

/**
 * The comment marks every piece of content in `[from, to]` carries: the
 * comments whose content typing over the range replaces. Anchors without
 * content (comment references, range boundaries) do not decide it.
 */
export const commentMarksCovering = (
  doc: ProseMirrorNode,
  from: number,
  to: number,
): readonly Mark[] => {
  let covering: readonly Mark[] | null = null;
  doc.nodesBetween(from, to, (node) => {
    if (!isCoverableContent(node)) return true;
    const comments = node.marks.filter(isCommentMark);
    covering =
      covering === null
        ? comments
        : covering.filter((mark) => mark.isInSet(comments));
    return false;
  });
  return covering ?? [];
};

/**
 * Where typing over a commented range left the caret, still inside the
 * comments it typed into. A comment mark does not extend at its end, so the
 * caret right after the typed text would otherwise type outside them.
 */
type TypingInComments = { comments: readonly Mark[]; at: number };

const typingInCommentsOwner = createPluginKeyOwner<TypingInComments | null>(
  "typingInComments",
);

/**
 * Keep typed text inside the comments it belongs to.
 *
 * Typing over a range keeps the comments that cover all of it on the typed
 * text, and the caret stays inside them: the next characters typed at the
 * caret join them, until the caret moves. Without this the text typed over a
 * comment's whole range lands outside it, and the comment loses its text.
 */
export const keepTypedTextInComments = ({
  state,
  tr,
  replaced,
  typed,
}: {
  /** The state the typing started from. */
  state: EditorState;
  tr: Transaction;
  /** The range the typing replaced, in `state.doc`. */
  replaced: { from: number; to: number };
  /** The typed text, in `tr.doc`. */
  typed: { from: number; to: number };
}): void => {
  if (typed.from >= typed.to) return;
  const carried = (() => {
    if (replaced.from < replaced.to)
      return commentMarksCovering(state.doc, replaced.from, replaced.to);
    const typing = typingInCommentsOwner.getState(state);
    if (!typing || typing.at !== replaced.from) return [];
    // Only comments the text before the caret still carries: one deleted
    // since must not come back.
    const before = state.doc.resolve(replaced.from).nodeBefore?.marks ?? [];
    return typing.comments.filter((mark) => mark.isInSet(before));
  })();
  if (carried.length === 0) return;
  // Added in one order, so every typed character holds the same mark set
  // and the text stays one run.
  for (const mark of carried) tr.removeMark(typed.from, typed.to, mark);
  for (const mark of carried) tr.addMark(typed.from, typed.to, mark);
  tr.setMeta(typingInCommentsOwner.key, { comments: carried, at: typed.to });
};

export const dispatchEditorTextInput = <
  TView extends TextInputDispatchTarget<TView>,
>(
  view: TView,
  input: string | TextInput,
): void => {
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
      const typesOverSelection =
        !selection.empty && selection.from === from && selection.to === to;
      // Selected cells or a selected block end between blocks, where the text
      // needs a paragraph. The selection replaces itself, as it does for a key
      // ProseMirror handles, and leaves the caret after the text.
      if (typesOverSelection && !selection.$from.parent.inlineContent) {
        return tr.insertText(text);
      }
      tr.insertText(text, from, to);
      // Typing over a selection leaves the caret after the typed text;
      // `insertText` puts it at the selection's mapped end, which for a node
      // selection (an image, a field) is before the text.
      if (typesOverSelection)
        tr.setSelection(TextSelection.create(tr.doc, from + text.length));
      keepTypedTextInComments({
        state: view.state,
        tr,
        replaced: { from, to },
        typed: { from, to: from + text.length },
      });
      return tr;
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
    tr.setSelection(TextSelection.create(tr.doc, start.from + text.length));
    keepTypedTextInComments({
      state: view.state,
      tr,
      replaced: { from: start.from, to: end.to },
      typed: { from: start.from, to: start.from + text.length },
    });
    return tr;
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
): {
  from: number;
  to: number;
} | null => {
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
export const createTextInputPlugin = (): Plugin => {
  // PM can clear view.composing before its final DOM flush. Keep that commit
  // native through the end-of-composition microtask, just like suggestion mode.
  // Key by view because a plugin can be shared by multiple editor states/views.
  const compositions = new WeakMap<EditorView, Event>();
  return new Plugin<TypingInComments | null>({
    key: typingInCommentsOwner.createKey(),
    state: {
      init: () => null,
      apply: (tr, value) => {
        const typed: TypingInComments | null | undefined = tr.getMeta(
          typingInCommentsOwner.key,
        );
        if (typed !== undefined) return typed;
        if (!value) return null;
        // Kept only while the caret stays where the typing left it. Anything
        // put in at that point (a restored reference) goes after it.
        const at = tr.mapping.map(value.at, -1);
        return tr.selection.empty && tr.selection.from === at
          ? { comments: value.comments, at }
          : null;
      },
    },
    // Text entered at that caret by a route other than the input handlers
    // (a composition the browser commits, a native insertion) joins the same
    // comments, as typed characters do.
    appendTransaction: (transactions, oldState, newState) => {
      const typing = typingInCommentsOwner.getState(oldState);
      if (
        !typing ||
        !transactions.some((tr) => tr.docChanged) ||
        transactions.some(
          (tr) => tr.getMeta(typingInCommentsOwner.key) !== undefined,
        )
      )
        return null;
      const mapping = new Mapping();
      for (const tr of transactions) mapping.appendMapping(tr.mapping);
      const from = mapping.map(typing.at, -1);
      const to = mapping.map(typing.at, 1);
      const { selection, doc } = newState;
      if (from >= to || !selection.empty || selection.from !== to) return null;
      const $from = doc.resolve(from);
      if (!$from.parent.inlineContent || !$from.sameParent(doc.resolve(to)))
        return null;
      let textOnly = true;
      doc.nodesBetween(from, to, (node) => {
        if (node.isInline && !node.isText) textOnly = false;
        return textOnly;
      });
      if (!textOnly) return null;
      const before = oldState.doc.resolve(typing.at).nodeBefore?.marks ?? [];
      const carried = typing.comments.filter((mark) => mark.isInSet(before));
      if (carried.length === 0) return null;
      const tr = newState.tr;
      for (const mark of carried) tr.removeMark(from, to, mark);
      for (const mark of carried) tr.addMark(from, to, mark);
      return tr.setMeta(typingInCommentsOwner.key, {
        comments: carried,
        at: to,
      });
    },
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
        keydown(view, event) {
          if (!compositions.has(view) && !isCompositionKeyDown(event))
            adoptNativeSelection(view);
          return false;
        },
        beforeinput(view, event) {
          if (compositions.has(view)) return false;
          adoptNativeSelection(view);
          return handleEditorBeforeInput(view, event);
        },
      },
    },
  });
};
