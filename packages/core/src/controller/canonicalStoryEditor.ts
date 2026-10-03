import type { OpStory } from "@stll/docx-core/ops";
import type { Transaction } from "prosemirror-state";
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
  onRefusal: ((reason: string) => void) | undefined;
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
  const refuse = (message: string) => {
    if (onRefusal) onRefusal(message);
    else throw new CanonicalSessionError({ message });
  };
  const history = (direction: "undo" | "redo") => {
    const view = getView();
    const accepted = getApi()?.applyCanonicalStoryHistory({ view, story, direction }) ?? false;
    if (accepted && !view.isDestroyed) onSelectionChange();
    return accepted;
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
        beforeinput: (view: EditorView, event: InputEvent) =>
          enabled() && boundary.handleDOMEvents.beforeinput(view, event),
        mousedown: (view: EditorView) => enabled() && boundary.handleDOMEvents.mousedown(view),
        compositionstart: (view: EditorView) =>
          enabled() && boundary.handleDOMEvents.compositionstart(view),
        compositionend: (view: EditorView) =>
          enabled() && boundary.handleDOMEvents.compositionend(view),
        input: (view: EditorView) => enabled() && boundary.handleDOMEvents.input(view),
        blur: (view: EditorView) => enabled() && boundary.handleDOMEvents.blur(view),
        paste: (view: EditorView, event: Event) =>
          enabled() && boundary.handleDOMEvents.paste(view, event),
        cut: (view: EditorView, event: Event) =>
          enabled() && boundary.handleDOMEvents.cut(view, event),
        drop: (view: EditorView, event: Event) =>
          enabled() && boundary.handleDOMEvents.drop(view, event),
      },
    },
  };
};
