import { Result } from "better-result";

import type { EditorState, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { CanonicalSessionError } from "./canonicalSession";
import { splitsSurrogatePair } from "../ai-edits/character-boundaries";

const COMPOSITION_FLUSH_DELAY_MS = 25;

type CompositionOptions = {
  begin: () => boolean;
  end: () => void;
  replace: (input: { from: number; to: number; text: string; semantic: "composition" }) => void;
  refuse: (reason: string) => void;
};

type CompositionState =
  | { type: "committed" }
  | { type: "provisional"; baseline: EditorState; phase: "active" | "ended" };

/** Native IME owns the provisional view; the captured canonical projection stays unchanged. */
export const createCanonicalComposition = (options: CompositionOptions) => {
  let state: CompositionState = { type: "committed" };
  let authorizedNativeState: EditorState | null = null;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const clearTimer = () => {
    if (timer !== undefined) clearTimeout(timer);
    timer = undefined;
  };
  const replacement = (baseline: EditorState, proposed: EditorState) => {
    const { $from, $to } = baseline.selection;
    if (!$from.sameParent($to) || !$from.parent.isTextblock) return null;
    const start = $from.start();
    const next = proposed.doc.nodeAt(start - 1);
    if (!next || next.type !== $from.parent.type || next.content.size !== next.textContent.length)
      return null;
    const before = $from.parent.textContent;
    const after = next.textContent;
    let from = 0;
    while (from < before.length && from < after.length && before[from] === after[from]) from++;
    while (splitsSurrogatePair(before, from) || splitsSurrogatePair(after, from)) from--;
    let oldEnd = before.length;
    let newEnd = after.length;
    while (oldEnd > from && newEnd > from && before[oldEnd - 1] === after[newEnd - 1]) {
      oldEnd--;
      newEnd--;
    }
    while (splitsSurrogatePair(before, oldEnd) || splitsSurrogatePair(after, newEnd)) {
      oldEnd++;
      newEnd++;
    }
    const text = after.slice(from, newEnd);
    const expected = baseline.tr.insertText(text, start + from, start + oldEnd);
    if (!expected.doc.eq(proposed.doc)) return null;
    return { from: start + from, to: start + oldEnd, text, semantic: "composition" as const };
  };
  const finish = (view: EditorView, commit: boolean) => {
    clearTimer();
    authorizedNativeState = null;
    const pending = state;
    if (pending.type !== "provisional") return;
    const input = commit ? replacement(pending.baseline, view.state) : null;
    state = { type: "committed" };
    if (!view.isDestroyed) view.updateState(pending.baseline);
    options.end();
    if (!commit || view.isDestroyed) return;
    if (input?.from === input?.to && input?.text === "") return;
    if (!input) {
      options.refuse(
        "Composition changed unsupported content; the canonical document was restored.",
      );
      return;
    }
    options.replace(input);
  };
  const schedule = (view: EditorView) => {
    clearTimer();
    // Allow the native final input and PM's delayed composition flush to settle.
    timer = setTimeout(() => finish(view, true), COMPOSITION_FLUSH_DELAY_MS);
  };
  return {
    get active() {
      return state.type === "provisional";
    },
    start: (view: EditorView) => {
      if (state.type === "provisional") return;
      if (!options.begin()) return;
      state = { type: "provisional", baseline: view.state, phase: "active" };
    },
    ended: (view: EditorView) => {
      if (state.type !== "provisional") return;
      state = { type: "provisional", baseline: state.baseline, phase: "ended" };
      schedule(view);
    },
    authorizeNative: (view: EditorView) => {
      authorizedNativeState = view.state;
    },
    flushed: (view: EditorView) => {
      const authorized = authorizedNativeState;
      queueMicrotask(() => {
        if (authorizedNativeState === authorized) authorizedNativeState = null;
      });
      if (state.type === "provisional" && state.phase === "ended") schedule(view);
    },
    recover: (view: EditorView) => finish(view, true),
    cancel: (view: EditorView) => finish(view, false),
    reset: () => {
      clearTimer();
      authorizedNativeState = null;
      if (state.type === "provisional") options.end();
      state = { type: "committed" };
    },
    accept: (view: EditorView, transaction: Transaction) => {
      if (state.type !== "provisional") return false;
      if (
        typeof transaction.getMeta("composition") !== "number" &&
        authorizedNativeState !== view.state
      ) {
        options.refuse("Other edits must wait until composition finishes.");
        return true;
      }
      authorizedNativeState = null;
      const applied = Result.try({
        try: () => view.state.applyTransaction(transaction),
        catch: (cause) =>
          new CanonicalSessionError({ message: `Composition projection failed: ${String(cause)}` }),
      });
      if (applied.isErr()) {
        finish(view, false);
        options.refuse(applied.error.message);
        return true;
      }
      // Validate the entire proposal against one plain replacement, including plugin output.
      if (!replacement(state.baseline, applied.value.state)) {
        finish(view, false);
        options.refuse(
          "Composition changed unsupported content; the canonical document was restored.",
        );
        return true;
      }
      view.updateState(applied.value.state);
      return true;
    },
  };
};
