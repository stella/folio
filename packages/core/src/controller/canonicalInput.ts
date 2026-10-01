import type { EditorState, Transaction } from "prosemirror-state";
import { Selection, TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { ReplaceStep } from "prosemirror-transform";

import { handleEditorBeforeInput } from "../prosemirror/textInput";
import { deletionRange } from "./canonicalSession";

type ReplaceTextInput = { from: number; to: number; text: string };

type CanonicalInputOptions = {
  replace: (input: ReplaceTextInput) => void;
  undo: () => boolean;
  redo: () => boolean;
  refuse: (reason: string) => void;
};

type NativeProposal =
  | { type: "idle" }
  | { type: "composition"; phase: "active" | "ended" }
  | { type: "refused" }
  | { type: "replacement"; state: EditorState; input: ReplaceTextInput };

/** Native changes require a preceding, classified input event at the same state. */
export const createCanonicalInputBoundary = (options: CanonicalInputOptions) => {
  let proposal: NativeProposal = { type: "idle" };
  let authorizedInput: { view: EditorView; state: EditorState } | null = null;
  const repaint = (view: EditorView) => {
    queueMicrotask(() => {
      if (!view.isDestroyed) view.updateState(view.state);
    });
  };
  const refuseEvent = (event: Event, reason: string) => {
    event.preventDefault();
    if (proposal.type !== "composition" && proposal.type !== "refused") options.refuse(reason);
    if (proposal.type !== "composition") proposal = { type: "refused" };
    return true;
  };

  const refuseNativeMutation = (view: EditorView) => {
    if (proposal.type !== "composition" && proposal.type !== "refused") {
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
    },
    handleTextInput: (view: EditorView, from: number, to: number, text: string) => {
      const pending = proposal;
      if (authorizedInput?.view === view && authorizedInput.state === view.state) {
        options.replace({ from, to, text });
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
      const modifier = event.metaKey || event.ctrlKey;
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
        return refuseEvent(event, "Paragraph structure edits are unavailable in this session.");
      }
      if (event.key !== "Backspace" && event.key !== "Delete") return false;
      event.preventDefault();
      if (proposal.type === "composition" && proposal.phase === "ended" && !view.composing) {
        proposal = { type: "idle" };
        repaint(view);
      }
      if (modifier || event.altKey || view.composing || proposal.type === "composition") {
        if (proposal.type !== "composition")
          options.refuse("Only plain character deletion is available in this session.");
        return true;
      }
      const range = deletionRange(view.state, event.key === "Backspace" ? "backward" : "forward");
      if (range.isErr()) options.refuse(range.error.message);
      else options.replace({ ...range.value, text: "" });
      return true;
    },
    handleDOMEvents: {
      compositionstart: (_view: EditorView, event: Event) => {
        if (proposal.type !== "composition")
          options.refuse("Composition is unavailable in the experimental canonical session.");
        proposal = { type: "composition", phase: "active" };
        event.preventDefault();
        return true;
      },
      compositionend: (view: EditorView) => {
        if (proposal.type === "composition") proposal = { type: "composition", phase: "ended" };
        const composition = proposal;
        queueMicrotask(() => {
          if (view.isDestroyed || proposal !== composition) return;
          view.updateState(view.state);
        });
        return false;
      },
      blur: (view: EditorView) => {
        proposal = { type: "idle" };
        repaint(view);
        return false;
      },
      beforeinput: (view: EditorView, event: InputEvent) => {
        if (
          view.composing ||
          event.isComposing ||
          event.inputType === "insertCompositionText" ||
          event.inputType === "insertFromComposition" ||
          event.inputType === "deleteCompositionText" ||
          event.inputType === "deleteByComposition"
        ) {
          return refuseEvent(event, "Composition is unavailable in this session.");
        }
        // A new, non-composition event recovers an IME missing compositionend.
        if (proposal.type === "composition") {
          proposal = { type: "idle" };
          repaint(view);
        }
        if (event.inputType === "insertText" || event.inputType === "insertReplacementText") {
          if (event.cancelable) {
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
            },
          };
          return false;
        }
        if (
          event.inputType === "deleteContentBackward" ||
          event.inputType === "deleteContentForward"
        ) {
          const range = deletionRange(
            view.state,
            event.inputType === "deleteContentBackward" ? "backward" : "forward",
          );
          if (range.isErr()) return refuseEvent(event, range.error.message);
          const input = { ...range.value, text: "" };
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
        const pending = proposal;
        queueMicrotask(() => {
          if (view.isDestroyed) return;
          view.updateState(view.state);
          if (proposal === pending && pending.type !== "composition") proposal = { type: "idle" };
        });
        return false;
      },
      paste: (_view: EditorView, event: Event) =>
        refuseEvent(event, "Paste is unavailable in this session."),
      cut: (_view: EditorView, event: Event) =>
        refuseEvent(event, "Cut is unavailable in this session."),
      drop: (_view: EditorView, event: Event) =>
        refuseEvent(event, "Drop is unavailable in this session."),
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
