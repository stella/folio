import { TextSelection, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import type { OpStory } from "@stll/docx-core/ops";

import { CanonicalSessionError } from "../canonical/canonicalSession";
import {
  adoptNativeSelection,
  isCompositionKeyDown,
} from "../prosemirror/nativeSelection";
import {
  CANONICAL_GAP,
  type CanonicalRefusalGround,
} from "../types/canonicalCapabilities";
import {
  createCanonicalInputBoundary,
  unclassifiedMutationMessage,
} from "./canonicalInput";
import type { HiddenEditorApi } from "./hiddenEditorApi";

type CanonicalStoryEditorOptions = {
  story: OpStory;
  getView: () => EditorView;
  getApi: () => HiddenEditorApi | null;
  enabled: () => boolean;
  onSelectionChange: () => void;
  onRefusal:
    ((reason: string, ground: CanonicalRefusalGround) => void) | undefined;
};

/**
 * Show the story's current projection in its open editor, with the selection
 * the session holds for the story: the one the editor last reported, moved
 * over every edit since. A peer's edit can move it without changing what the
 * story shows, so it is restored then too. An editor that never reported one
 * keeps its position, within the new content.
 */
export const refreshCanonicalStoryView = ({
  view,
  story,
  api,
}: {
  view: EditorView;
  story: OpStory;
  api: HiddenEditorApi | null | undefined;
}): void => {
  const projected = api?.getCanonicalStoryProjection(story);
  if (!api || !projected) return;
  const held = api.getCanonicalStorySelection(story);
  if (view.state.doc.eq(projected)) {
    const { anchor, head } = view.state.selection;
    if (!held || (held.anchor === anchor && held.head === head)) return;
    const moved = view.state.tr.setSelection(
      TextSelection.create(view.state.doc, held.anchor, held.head),
    );
    moved.setMeta("addToHistory", false);
    view.updateState(view.state.apply(moved));
    return;
  }
  const transaction = view.state.tr.replaceWith(
    0,
    view.state.doc.content.size,
    projected.content,
  );
  transaction.setSelection(
    held
      ? TextSelection.create(transaction.doc, held.anchor, held.head)
      : TextSelection.near(
          transaction.doc.resolve(
            Math.min(view.state.selection.anchor, projected.content.size),
          ),
        ),
  );
  transaction.setMeta("addToHistory", false);
  view.updateState(view.state.apply(transaction));
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
    const ground = { gap: CANONICAL_GAP.dispatch };
    const message = reason;
    if (onRefusal) onRefusal(message, ground);
    else
      throw new CanonicalSessionError({ ground, message, reason: "refused" });
  };
  // Only the undo keys reach this; the host's own undo goes through the API.
  const history = (direction: "undo" | "redo") => {
    const view = getView();
    const accepted =
      getApi()?.applyCanonicalStoryHistory({
        view,
        story,
        direction,
        shortcut: true,
      }) ?? false;
    if (accepted && !view.isDestroyed) onSelectionChange();
    return accepted;
  };
  const syncNativeSelection = (view: EditorView) => {
    if (!boundary.isComposing) adoptNativeSelection(view);
  };
  const boundary = createCanonicalInputBoundary({
    breakUndoGroup: () => {
      getApi()?.updateCanonicalInputLifecycle("breakUndoGroup");
    },
    beginComposition: () =>
      getApi()?.updateCanonicalInputLifecycle("beginComposition") ?? false,
    endComposition: () => {
      getApi()?.updateCanonicalInputLifecycle("endComposition");
    },
    replace: (intent) => {
      const view = getView();
      if (
        getApi()?.replaceCanonicalStoryText({ view, story, intent }) &&
        !view.isDestroyed
      )
        onSelectionChange();
    },
    undo: () => history("undo"),
    redo: () => history("redo"),
    refuse,
  });
  return {
    dispatch: (transaction: Transaction): boolean => {
      if (!enabled()) return false;
      const view = getView();
      // Public operations on this story work through the editor in use.
      getApi()?.noteCanonicalStoryView({ story, view });
      if (transaction.docChanged) {
        if (boundary.acceptComposition(view, transaction)) return true;
        if (!boundary.commitNativeProposal(view, transaction))
          boundary.refuseNativeMutation(view, transaction);
        view.updateState(view.state);
        return true;
      }
      if (transaction.selectionSet && !boundary.isComposing)
        getApi()?.updateCanonicalInputLifecycle("breakUndoGroup");
      const applied = view.state.applyTransaction(transaction);
      if (!applied.state.doc.eq(view.state.doc)) {
        refuse(
          unclassifiedMutationMessage(
            applied.transactions,
            "A plugin attempted an unclassified canonical story mutation.",
          ),
        );
        view.updateState(view.state);
        return true;
      }
      view.updateState(applied.state);
      if (applied.transactions.some((step) => step.selectionSet)) {
        // The selection it now shows, which a peer's edit moves with the text.
        getApi()?.noteCanonicalStoryView({ story, view });
        onSelectionChange();
      }
      return true;
    },
    props: {
      handleKeyDown: (view: EditorView, event: KeyboardEvent): boolean =>
        enabled() && boundary.handleKeyDown(view, event),
      handleTextInput: (
        view: EditorView,
        from: number,
        to: number,
        text: string,
      ): boolean => enabled() && boundary.handleTextInput(view, from, to, text),
      handleKeyPress: (view: EditorView, event: KeyboardEvent): boolean =>
        enabled() && boundary.handleKeyPress(view, event),
      handleDOMEvents: {
        keydown: (view: EditorView, event: KeyboardEvent): boolean => {
          if (!enabled()) return false;
          if (!isCompositionKeyDown(event)) syncNativeSelection(view);
          return boundary.handleDOMEvents.keydown(view, event);
        },
        // Native navigation completes before keyup; synchronize it before typing
        // rather than waiting for the asynchronous selection observer.
        keyup: (view: EditorView, event: KeyboardEvent) => {
          if (enabled() && /^(?:Arrow|Home$|End$|Page)/u.test(event.key))
            syncNativeSelection(view);
          return false;
        },
        mouseup: (view: EditorView) => {
          if (enabled()) syncNativeSelection(view);
          return false;
        },
        beforeinput: (view: EditorView, event: InputEvent): boolean => {
          if (!enabled()) return false;
          if (!event.isComposing) syncNativeSelection(view);
          return boundary.handleDOMEvents.beforeinput(view, event);
        },
        mousedown: (view: EditorView): boolean =>
          enabled() && boundary.handleDOMEvents.mousedown(view),
        compositionstart: (view: EditorView): boolean =>
          enabled() && boundary.handleDOMEvents.compositionstart(view),
        compositionend: (view: EditorView, event: Event): boolean =>
          enabled() && boundary.handleDOMEvents.compositionend(view, event),
        input: (view: EditorView): boolean =>
          enabled() && boundary.handleDOMEvents.input(view),
        blur: (view: EditorView): boolean =>
          enabled() && boundary.handleDOMEvents.blur(view),
        paste: (view: EditorView, event: ClipboardEvent): boolean =>
          enabled() && boundary.handleDOMEvents.paste(view, event),
        cut: (view: EditorView, event: ClipboardEvent): boolean =>
          enabled() && boundary.handleDOMEvents.cut(view, event),
        drop: (view: EditorView, event: Event): boolean =>
          enabled() && boundary.handleDOMEvents.drop(view, event),
      },
    },
  };
};
