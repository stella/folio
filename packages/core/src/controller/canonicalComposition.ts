import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { Result } from "better-result";

import type { EditorState, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { CanonicalSessionError } from "./canonicalSession";
import { setCanonicalInputTimer } from "./canonicalInputTimer";
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
  | {
      type: "provisional" | "refused";
      baseline: EditorState;
      view: EditorView;
      phase: "active" | "ended";
    };

type CompositionFinish =
  | { type: "idle" | "cancelled" | "refused" }
  | { type: "committed"; text: string | null };

/** Native IME owns the provisional view; the captured canonical projection stays unchanged. */
export const createCanonicalComposition = (options: CompositionOptions) => {
  let state: CompositionState = { type: "committed" };
  let authorizedNativeState: EditorState | null = null;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const clearTimer = () => {
    if (timer !== undefined) clearTimeout(timer);
    timer = undefined;
  };
  const selectedReplacement = (baseline: EditorState, proposed: EditorState) => {
    const { from, to } = baseline.selection;
    const size = proposed.doc.content.size - baseline.doc.content.size + to - from;
    if (size < 0 || from + size > proposed.doc.content.size) return null;
    const text = proposed.doc.textBetween(from, from + size, "", "");
    if (text.length !== size || !baseline.tr.insertText(text, from, to).doc.eq(proposed.doc))
      return null;
    return { from, to, text, semantic: "composition" as const };
  };
  const replacement = (baseline: EditorState, proposed: EditorState) => {
    const { $from, $to } = baseline.selection;
    if (!$from.parent.isTextblock || !$to.parent.isTextblock) return null;
    // Native replacement may drop non-inclusive run metadata across the whole
    // captured selection even when its text shares a prefix or suffix.
    const selected = selectedReplacement(baseline, proposed);
    if (selected) {
      if (
        $from.sameParent($to) &&
        selected.text === baseline.doc.textBetween(selected.from, selected.to, "", "")
      ) {
        return {
          from: selected.from,
          to: selected.from,
          text: "",
          semantic: "composition" as const,
        };
      }
      return selected;
    }
    // A selected paragraph boundary can disappear only through its exact replacement.
    if (!$from.sameParent($to)) return null;
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
  const finish = (view: EditorView, commit: boolean): CompositionFinish => {
    clearTimer();
    authorizedNativeState = null;
    const pending = state;
    if (pending.type === "committed") return { type: "idle" };
    const input =
      commit && pending.type === "provisional" ? replacement(pending.baseline, view.state) : null;
    // A final beforeinput carries the whole selected replacement, even when
    // the native observer produced a smaller diff with a shared prefix/suffix.
    const { from, to } = pending.baseline.selection;
    const size = view.state.doc.content.size - pending.baseline.doc.content.size + to - from;
    const text =
      size >= 0 && from + size <= view.state.doc.content.size
        ? view.state.doc.textBetween(from, from + size, "", "")
        : null;
    // Acceptance is validated by replacement(); payload identity depends on text,
    // including when the accepted minimal diff retains non-inclusive metadata.
    const committedText = text !== null && text.length === size ? text : null;
    state = { type: "committed" };
    if (!view.isDestroyed) {
      // Recovery may finish without a native end event (for example a refused
      // final beforeinput). Let PM end its own lifecycle before restoring DOM.
      if (view.composing) view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
      view.updateState(pending.baseline);
    }
    options.end();
    if (!commit || view.isDestroyed) return { type: "cancelled" };
    if (pending.type === "refused") return { type: "refused" };
    if (input?.from === input?.to && input?.text === "")
      return { type: "committed", text: committedText };
    if (!input) {
      options.refuse(
        "Composition changed unsupported content; the canonical document was restored.",
      );
      return { type: "refused" };
    }
    options.replace(input);
    return { type: "committed", text: committedText };
  };
  const schedule = (view: EditorView) => {
    clearTimer();
    // Allow the native final input and PM's delayed composition flush to settle.
    timer = setCanonicalInputTimer(() => finish(view, true), COMPOSITION_FLUSH_DELAY_MS);
  };
  return {
    // Keep declaration emission tied to the authored union, not inferred member order.
    get status(): CompositionState["type"] {
      return state.type;
    },
    get active() {
      return state.type !== "committed";
    },
    start: (view: EditorView) => {
      if (state.type !== "committed") return;
      if (!options.begin()) return;
      state = { type: "provisional", baseline: view.state, view, phase: "active" };
    },
    ended: (view: EditorView) => {
      if (state.type === "committed") return;
      state = { type: state.type, baseline: state.baseline, view: state.view, phase: "ended" };
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
      if (state.type !== "committed" && state.phase === "ended") schedule(view);
    },
    recover: (view: EditorView) => finish(view, true),
    cancel: (view: EditorView) => finish(view, false),
    reset: () => {
      if (state.type !== "committed") finish(state.view, false);
      else {
        clearTimer();
        authorizedNativeState = null;
      }
    },
    accept: (view: EditorView, transaction: Transaction) => {
      if (state.type === "committed") return false;
      if (state.type === "refused") {
        view.updateState(state.baseline);
        return true;
      }
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
          new CanonicalSessionError({
            gap: CANONICAL_GAP.dispatch,
            reason: "refused",
            message: `Composition projection failed: ${String(cause)}`,
          }),
      });
      if (applied.isErr()) {
        view.updateState(state.baseline);
        state = { type: "refused", baseline: state.baseline, view: state.view, phase: state.phase };
        options.refuse(applied.error.message);
        return true;
      }
      // Validate the entire proposal against one plain replacement, including plugin output.
      if (!replacement(state.baseline, applied.value.state)) {
        view.updateState(state.baseline);
        state = { type: "refused", baseline: state.baseline, view: state.view, phase: state.phase };
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
