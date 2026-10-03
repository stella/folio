import type { EditorState, Transaction } from "prosemirror-state";
import { Selection, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { ReplaceStep } from "prosemirror-transform";

import { handleEditorBeforeInput } from "../prosemirror/textInput";
import { createCanonicalComposition } from "./canonicalComposition";
import { deletionRange, isCanonicalJoinBoundary } from "./canonicalSession";

type ReplaceTextInput = {
  from: number;
  to: number;
  text: string;
  semantic?: "typing" | "deleteBackward" | "deleteForward" | "composition";
};

type CanonicalInputOptions = {
  replace: (input: ReplaceTextInput) => void;
  breakUndoGroup?: () => void;
  beginComposition?: () => boolean;
  endComposition?: () => void;
  split?: () => void;
  join?: (direction: "backward" | "forward") => void;
  undo: () => boolean;
  redo: () => boolean;
  refuse: (reason: string) => void;
};

type NativeProposal =
  | { type: "idle" }
  | { type: "refused" }
  | { type: "replacement"; state: EditorState; input: ReplaceTextInput };

/** Native changes require a preceding, classified input event at the same state. */
export const createCanonicalInputBoundary = (options: CanonicalInputOptions) => {
  const composition = createCanonicalComposition({
    begin: () => options.beginComposition?.() ?? true,
    end: () => options.endComposition?.(),
    replace: options.replace,
    refuse: options.refuse,
  });
  let proposal: NativeProposal = { type: "idle" };
  let authorizedInput: { view: EditorView; state: EditorState } | null = null;
  const beginGesture = () => {
    proposal = { type: "idle" };
    authorizedInput = null;
  };
  const closeGroup = () => options.breakUndoGroup?.();
  const repaint = (view: EditorView) => {
    queueMicrotask(() => {
      if (!view.isDestroyed) view.updateState(view.state);
    });
  };
  const refuseEvent = (event: Event, reason: string) => {
    event.preventDefault();
    closeGroup();
    if (!composition.active && proposal.type !== "refused") options.refuse(reason);
    if (!composition.active) proposal = { type: "refused" };
    return true;
  };

  const refuseNativeMutation = (view: EditorView) => {
    if (!composition.active && proposal.type !== "refused") {
      options.refuse("Unclassified native text is unavailable in this session.");
      proposal = { type: "refused" };
    }
    repaint(view);
  };

  const takeNativeProposal = (
    state: EditorState,
    transaction: Transaction,
  ): ReplaceTextInput | null => {
    const pending = proposal;
    if (pending.type !== "replacement") return null;
    proposal = { type: "idle" };
    if (pending.state !== state || transaction.steps.length !== 1) return null;
    const step = transaction.steps.at(0);
    if (
      !(step instanceof ReplaceStep) ||
      ("structure" in step && step.structure === true) ||
      step.from !== pending.input.from ||
      step.to !== pending.input.to
    )
      return null;
    if (step.slice.openStart !== 0 || step.slice.openEnd !== 0) return null;
    let text = "";
    let plain = true;
    step.slice.content.forEach((node) => {
      if (!node.isText) plain = false;
      text += node.text ?? "";
    });
    if (!plain || text !== pending.input.text) return null;
    // Marks, attrs and unrelated nodes cannot hitch a ride on classified input.
    const expected = state.tr.insertText(text, step.from, step.to);
    return expected.doc.eq(transaction.doc) ? pending.input : null;
  };

  return {
    reset: () => {
      proposal = { type: "idle" };
      authorizedInput = null;
      composition.reset();
    },
    get isComposing() {
      return composition.active;
    },
    acceptComposition: composition.accept,
    handleTextInput: (view: EditorView, from: number, to: number, text: string) => {
      if (composition.active) return false;
      const pending = proposal;
      if (authorizedInput?.view === view && authorizedInput.state === view.state) {
        options.replace({ from, to, text, semantic: "typing" });
        return true;
      }
      if (
        pending.type === "replacement" &&
        pending.state === view.state &&
        pending.input.from === from &&
        pending.input.to === to &&
        pending.input.text === text
      ) {
        proposal = { type: "idle" };
        options.replace(pending.input);
        return true;
      }
      refuseNativeMutation(view);
      return true;
    },
    handleKeyDown: (view: EditorView, event: KeyboardEvent) => {
      beginGesture();
      if (composition.active) {
        if (event.key === "Escape") {
          event.preventDefault();
          composition.cancel(view);
          return true;
        }
        if (event.metaKey || event.ctrlKey) {
          event.preventDefault();
          return true;
        }
        return false;
      }
      const modifier = event.metaKey || event.ctrlKey;
      if (modifier || event.altKey || /^(?:Arrow|Home$|End$|Page|Escape$|Tab$)/u.test(event.key))
        closeGroup();
      if (modifier && event.key.toLowerCase() === "z") {
        event.preventDefault();
        if (event.shiftKey) options.redo();
        else options.undo();
        return true;
      }
      if (modifier && event.key.toLowerCase() === "y") {
        event.preventDefault();
        options.redo();
        return true;
      }
      if (modifier && event.key.toLowerCase() === "a") {
        event.preventDefault();
        view.dispatch(
          view.state.tr.setSelection(
            TextSelection.create(
              view.state.doc,
              Selection.atStart(view.state.doc).from,
              Selection.atEnd(view.state.doc).to,
            ),
          ),
        );
        return true;
      }
      if (event.key === "Enter") {
        if (!options.split || view.composing || composition.active)
          return refuseEvent(
            event,
            "Paragraph structure edits are unavailable during composition.",
          );
        event.preventDefault();
        options.split();
        return true;
      }
      if (event.key !== "Backspace" && event.key !== "Delete") return false;
      event.preventDefault();
      if (modifier || event.altKey || view.composing || composition.active) {
        closeGroup();
        if (!composition.active)
          options.refuse("Only plain character deletion is available in this session.");
        return true;
      }
      const direction = event.key === "Backspace" ? "backward" : "forward";
      const { selection } = view.state;
      if (options.join && selection.empty && isCanonicalJoinBoundary(view.state, direction)) {
        options.join(direction);
        return true;
      }
      const range = deletionRange(view.state, direction);
      if (range.isErr()) {
        closeGroup();
        options.refuse(range.error.message);
      } else
        options.replace({
          ...range.value,
          text: "",
          semantic: direction === "backward" ? "deleteBackward" : "deleteForward",
        });
      return true;
    },
    handleDOMEvents: {
      compositionstart: (view: EditorView) => {
        beginGesture();
        composition.start(view);
        return false;
      },
      compositionend: (view: EditorView) => {
        composition.ended(view);
        return false;
      },
      mousedown: (view: EditorView) => {
        composition.recover(view);
        beginGesture();
        closeGroup();
        return false;
      },
      blur: (view: EditorView) => {
        composition.recover(view);
        beginGesture();
        closeGroup();
        repaint(view);
        return false;
      },
      beforeinput: (view: EditorView, event: InputEvent) => {
        beginGesture();
        if (
          event.isComposing ||
          event.inputType === "insertCompositionText" ||
          event.inputType === "insertFromComposition" ||
          event.inputType === "deleteCompositionText" ||
          event.inputType === "deleteByComposition"
        ) {
          if (composition.active) {
            composition.authorizeNative(view);
            return false;
          }
          return refuseEvent(event, "Composition has no captured canonical baseline.");
        }
        // A new, non-composition event recovers an IME missing compositionend.
        if (composition.active) {
          const refusedNativeCommit = composition.status === "refused" && view.composing;
          composition.recover(view);
          proposal = { type: "idle" };
          if (refusedNativeCommit) {
            event.preventDefault();
            return true;
          }
        }
        if (event.inputType === "insertText" || event.inputType === "insertReplacementText") {
          if (event.cancelable) {
            if (view.composing && event.inputType === "insertText" && event.data !== null) {
              event.preventDefault();
              proposal = { type: "idle" };
              options.replace({
                from: view.state.selection.from,
                to: view.state.selection.to,
                text: event.data,
                semantic: "typing",
              });
              return true;
            }
            proposal = { type: "idle" };
            authorizedInput = { view, state: view.state };
            try {
              if (handleEditorBeforeInput(view, event)) return true;
              return refuseEvent(event, "This native replacement cannot be addressed safely.");
            } finally {
              authorizedInput = null;
            }
          }
          if (event.inputType !== "insertText" || event.data === null) {
            return refuseEvent(event, "This native replacement cannot be addressed safely.");
          }
          proposal = {
            type: "replacement",
            state: view.state,
            input: {
              from: view.state.selection.from,
              to: view.state.selection.to,
              text: event.data,
              semantic: "typing",
            },
          };
          return false;
        }
        if (event.inputType === "insertParagraph" && event.cancelable && options.split) {
          event.preventDefault();
          options.split();
          return true;
        }
        if (
          event.inputType === "deleteContentBackward" ||
          event.inputType === "deleteContentForward"
        ) {
          const direction = event.inputType === "deleteContentBackward" ? "backward" : "forward";
          const { selection } = view.state;
          if (
            event.cancelable &&
            options.join &&
            selection.empty &&
            isCanonicalJoinBoundary(view.state, direction)
          ) {
            event.preventDefault();
            options.join(direction);
            return true;
          }
          const range = deletionRange(view.state, direction);
          if (range.isErr()) return refuseEvent(event, range.error.message);
          const input = {
            ...range.value,
            text: "",
            semantic:
              event.inputType === "deleteContentBackward"
                ? ("deleteBackward" as const)
                : ("deleteForward" as const),
          };
          if (event.cancelable) {
            event.preventDefault();
            options.replace(input);
            return true;
          }
          proposal = { type: "replacement", state: view.state, input };
          return false;
        }
        if (event.inputType === "historyUndo" || event.inputType === "historyRedo") {
          event.preventDefault();
          if (event.inputType === "historyUndo") options.undo();
          else options.redo();
          return true;
        }
        return refuseEvent(event, `Input ${event.inputType} is unavailable in this session.`);
      },
      // A native DOM proposal is valid only until this event's observer flush.
      input: (view: EditorView) => {
        if (composition.active) {
          composition.flushed(view);
          return false;
        }
        const pending = proposal;
        queueMicrotask(() => {
          if (view.isDestroyed) return;
          view.updateState(view.state);
          if (proposal === pending && !composition.active) proposal = { type: "idle" };
        });
        return false;
      },
      paste: (_view: EditorView, event: Event) => {
        beginGesture();
        return refuseEvent(event, "Paste is unavailable in this session.");
      },
      cut: (_view: EditorView, event: Event) => {
        beginGesture();
        return refuseEvent(event, "Cut is unavailable in this session.");
      },
      drop: (_view: EditorView, event: Event) => {
        beginGesture();
        return refuseEvent(event, "Drop is unavailable in this session.");
      },
    },
    /** Returns a classified intent only; the proposed PM document is never authoritative. */
    takeNativeProposal,
    refuseNativeMutation,
    commitNativeProposal: (view: EditorView, transaction: Transaction) => {
      const input = takeNativeProposal(view.state, transaction);
      if (!input) return false;
      options.replace(input);
      return true;
    },
  };
};
