import { CANONICAL_GAP, type CanonicalGap } from "../types/canonicalCapabilities";
import type { OpStory } from "@stll/docx-core/ops";
import { AllSelection, TextSelection, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { CanonicalSessionError } from "./canonicalSession";
import { createCanonicalInputBoundary } from "./canonicalInput";
import type { HiddenEditorApi } from "./hiddenEditorApi";

type CanonicalStoryEditorOptions = {
  story: OpStory;
  getView: () => EditorView;
  getApi: () => HiddenEditorApi | null;
  enabled: () => boolean;
  onSelectionChange: () => void;
  onRefusal: ((reason: string, gap: CanonicalGap) => void) | undefined;
};

/** Secondary views compile classified native input into the shared Document journal. */
export const createCanonicalStoryEditor = ({
  story,
  getView,
  getApi,
  enabled,
  onRefusal,
  onSelectionChange,
}: CanonicalStoryEditorOptions) => {
  const refuse = (reason: string) => {
    const gap = CANONICAL_GAP.dispatch;
    const message = reason;
    if (onRefusal) onRefusal(message, gap);
    else throw new CanonicalSessionError({ gap, message, reason: "refused" });
  };
  const history = (direction: "undo" | "redo") => {
    const view = getView();
    const accepted = getApi()?.applyCanonicalStoryHistory({ view, story, direction }) ?? false;
    if (accepted && !view.isDestroyed) onSelectionChange();
    return accepted;
  };
  const syncNativeSelection = (view: EditorView) => {
    if (boundary.isComposing || view.composing) return;
    const selection = view.dom.ownerDocument.getSelection();
    if (
      !selection?.anchorNode ||
      !selection.focusNode ||
      !view.dom.contains(selection.anchorNode) ||
      !view.dom.contains(selection.focusNode)
    )
      return;
    const anchor = view.posAtDOM(selection.anchorNode, selection.anchorOffset);
    const head = view.posAtDOM(selection.focusNode, selection.focusOffset);
    if (
      view.state.selection instanceof AllSelection &&
      Math.min(anchor, head) === 0 &&
      Math.max(anchor, head) === view.state.doc.content.size
    )
      return;
    if (anchor === view.state.selection.anchor && head === view.state.selection.head) return;
    if (
      !view.state.doc.resolve(anchor).parent.inlineContent ||
      !view.state.doc.resolve(head).parent.inlineContent
    )
      return;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, anchor, head)));
  };
  const boundary = createCanonicalInputBoundary({
    breakUndoGroup: () => {
      getApi()?.updateCanonicalInputLifecycle("breakUndoGroup");
    },
    beginComposition: () => getApi()?.updateCanonicalInputLifecycle("beginComposition") ?? false,
    endComposition: () => {
      getApi()?.updateCanonicalInputLifecycle("endComposition");
    },
    replace: (intent) => {
      const view = getView();
      if (getApi()?.replaceCanonicalStoryText({ view, story, intent }) && !view.isDestroyed)
        onSelectionChange();
    },
    undo: () => history("undo"),
    redo: () => history("redo"),
    refuse,
  });
  return {
    dispatch: (transaction: Transaction) => {
      if (!enabled()) return false;
      const view = getView();
      if (transaction.docChanged) {
        if (boundary.acceptComposition(view, transaction)) return true;
        if (!boundary.commitNativeProposal(view, transaction)) boundary.refuseNativeMutation(view);
        view.updateState(view.state);
        return true;
      }
      if (transaction.selectionSet && !boundary.isComposing)
        getApi()?.updateCanonicalInputLifecycle("breakUndoGroup");
      const applied = view.state.applyTransaction(transaction);
      if (!applied.state.doc.eq(view.state.doc)) {
        refuse("A plugin attempted an unclassified canonical story mutation.");
        view.updateState(view.state);
        return true;
      }
      view.updateState(applied.state);
      if (applied.transactions.some((step) => step.selectionSet)) onSelectionChange();
      return true;
    },
    props: {
      handleKeyDown: (view: EditorView, event: KeyboardEvent) =>
        enabled() && boundary.handleKeyDown(view, event),
      handleTextInput: (view: EditorView, from: number, to: number, text: string) =>
        enabled() && boundary.handleTextInput(view, from, to, text),
      handleDOMEvents: {
        keydown: (view: EditorView, event: KeyboardEvent) =>
          enabled() && boundary.handleDOMEvents.keydown(view, event),
        // Native navigation completes before keyup; synchronize it before typing
        // rather than waiting for the asynchronous selection observer.
        keyup: (view: EditorView, event: KeyboardEvent) => {
          if (enabled() && /^(?:Arrow|Home$|End$|Page)/u.test(event.key)) syncNativeSelection(view);
          return false;
        },
        mouseup: (view: EditorView) => {
          if (enabled()) syncNativeSelection(view);
          return false;
        },
        beforeinput: (view: EditorView, event: InputEvent) =>
          enabled() && boundary.handleDOMEvents.beforeinput(view, event),
        mousedown: (view: EditorView) => enabled() && boundary.handleDOMEvents.mousedown(view),
        compositionstart: (view: EditorView) =>
          enabled() && boundary.handleDOMEvents.compositionstart(view),
        compositionend: (view: EditorView, event: Event) =>
          enabled() && boundary.handleDOMEvents.compositionend(view, event),
        input: (view: EditorView) => enabled() && boundary.handleDOMEvents.input(view),
        blur: (view: EditorView) => enabled() && boundary.handleDOMEvents.blur(view),
        paste: (view: EditorView, event: ClipboardEvent) =>
          enabled() && boundary.handleDOMEvents.paste(view, event),
        cut: (view: EditorView, event: ClipboardEvent) =>
          enabled() && boundary.handleDOMEvents.cut(view, event),
        drop: (view: EditorView, event: Event) =>
          enabled() && boundary.handleDOMEvents.drop(view, event),
      },
    },
  };
};
