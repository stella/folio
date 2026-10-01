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
  | { type: "composition" }
  | { type: "replacement"; state: EditorState; input: ReplaceTextInput };

/** Native changes require a preceding, classified input event at the same state. */
export const createCanonicalInputBoundary = (options: CanonicalInputOptions) => {
  let proposal: NativeProposal = { type: "idle" };
  const refuseEvent = (event: Event, reason: string) => {
    event.preventDefault();
    options.refuse(reason);
    return true;
  };

  return {
    reset: () => {
      proposal = { type: "idle" };
    },
    handleTextInput: (_view: EditorView, from: number, to: number, text: string) => {
      if (proposal.type === "composition") {
        options.refuse("Composition is unavailable in the experimental canonical session.");
        return true;
      }
      options.replace({ from, to, text });
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
      if (modifier || event.altKey || view.composing || proposal.type === "composition") {
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
        proposal = { type: "composition" };
        return refuseEvent(
          event,
          "Composition is unavailable in the experimental canonical session.",
        );
      },
      compositionend: (view: EditorView) => {
        const composition = proposal;
        // Native IME owns the DOM until its final flush. Restore afterwards;
        // neither its provisional nor final transaction may enter the model.
        queueMicrotask(() => {
          if (view.isDestroyed || proposal !== composition) return;
          view.updateState(view.state);
          proposal = { type: "idle" };
        });
        return false;
      },
      beforeinput: (view: EditorView, event: InputEvent) => {
        if (view.composing || event.isComposing || proposal.type === "composition") {
          return refuseEvent(event, "Composition is unavailable in this session.");
        }
        if (event.inputType === "insertText" || event.inputType === "insertReplacementText") {
          if (event.cancelable) return handleEditorBeforeInput(view, event);
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
      input: () => {
        queueMicrotask(() => {
          if (proposal.type === "replacement") proposal = { type: "idle" };
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
    takeNativeProposal: (state: EditorState, transaction: Transaction): ReplaceTextInput | null => {
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
      // Compare the full expected proposal so marks, attrs and unrelated nodes
      // cannot hitch a ride on a classified character input.
      const expected = state.tr.insertText(text, step.from, step.to);
      return expected.doc.eq(transaction.doc) ? pending.input : null;
    },
  };
};
