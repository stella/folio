import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { Result } from "better-result";

import { EditorState, TextSelection, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { CanonicalSessionError } from "./canonicalSession";
import {
  CANONICAL_COMPOSITION_FLUSH_DELAY_MS,
  setCanonicalInputTimer,
} from "./canonicalInputTimer";
import { splitsSurrogatePair } from "../ai-edits/character-boundaries";

type CompositionOptions = {
  begin: () => boolean;
  end: () => void;
  replace: (input: {
    from: number;
    to: number;
    text: string;
    semantic: "composition";
    compositionPhase?: "correction";
  }) => void;
  refuse: (reason: string) => void;
};

type CompositionState =
  | { type: "committed" }
  | {
      type: "provisional";
      baseline: EditorState;
      view: EditorView;
      phase: "active" | "ended";
    }
  | {
      type: "refused";
      baseline: EditorState;
      view: EditorView;
      phase: "active" | "ended";
      text: string | null;
    };

type CompositionFinish =
  | { type: "idle" | "cancelled" }
  | { type: "committed" | "refused"; text: string | null };

type CompletedReceipt = {
  view: EditorView;
  state: EditorState;
  from: number;
  to: number;
  text: string;
  marks: EditorState["storedMarks"];
};

type CompletedComposition =
  | { type: "idle" }
  | { type: "completed"; receipt: CompletedReceipt }
  | { type: "authorized"; receipt: CompletedReceipt; text: string };

/** Native IME owns the provisional view; the captured canonical projection stays unchanged. */
export const createCanonicalComposition = (options: CompositionOptions) => {
  let state: CompositionState = { type: "committed" };
  let authorizedNativeState: EditorState | null = null;
  let completed: CompletedComposition = { type: "idle" };
  let timer: ReturnType<typeof setTimeout> | undefined;
  const clearTimer = () => {
    if (timer !== undefined) clearTimeout(timer);
    timer = undefined;
  };
  const selectedText = (baseline: EditorState, proposed: Pick<EditorState, "doc">) => {
    const { from, to } = baseline.selection;
    const size = proposed.doc.content.size - baseline.doc.content.size + to - from;
    if (size < 0 || from + size > proposed.doc.content.size) return null;
    const text = proposed.doc.textBetween(from, from + size, "", "");
    return text.length === size ? text : null;
  };
  const selectedReplacement = (baseline: EditorState, proposed: EditorState) => {
    const { from, to } = baseline.selection;
    const text = selectedText(baseline, proposed);
    if (text === null || !baseline.tr.insertText(text, from, to).doc.eq(proposed.doc)) return null;
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
    if (!next || next.type !== $from.parent.type) return null;
    const previous = $from.parent.content;
    const changedFrom = previous.findDiffStart(next.content);
    if (changedFrom === null) {
      if (!baseline.doc.eq(proposed.doc)) return null;
      return {
        from: baseline.selection.from,
        to: baseline.selection.from,
        text: "",
        semantic: "composition" as const,
      };
    }
    const changedEnd = previous.findDiffEnd(next.content);
    if (changedEnd === null) return null;
    let from = changedFrom;
    let oldEnd = changedEnd.a;
    let newEnd = changedEnd.b;
    const overlap = from - Math.min(oldEnd, newEnd);
    if (overlap > 0) {
      oldEnd += overlap;
      newEnd += overlap;
    }
    // Leaf placeholders keep text offsets aligned with PM positions while
    // untouched inline atoms remain part of the exact reconstruction below.
    const before = previous.textBetween(0, previous.size, "", "\uFFFC");
    const after = next.content.textBetween(0, next.content.size, "", "\uFFFC");
    // Node diffs also report mark-only changes (a native rewrite can drop run
    // identity). Narrow to the text change so the reconstruction refuses them
    // instead of committing a same-text replacement.
    while (from < oldEnd && from < newEnd && before[from] === after[from]) from++;
    while (oldEnd > from && newEnd > from && before[oldEnd - 1] === after[newEnd - 1]) {
      oldEnd--;
      newEnd--;
    }
    while (splitsSurrogatePair(before, from) || splitsSurrogatePair(after, from)) from--;
    while (splitsSurrogatePair(before, oldEnd) || splitsSurrogatePair(after, newEnd)) {
      oldEnd++;
      newEnd++;
    }
    const removed = previous.cut(from, oldEnd);
    const inserted = next.content.cut(from, newEnd);
    const text = inserted.textBetween(0, inserted.size, "", "");
    if (
      removed.textBetween(0, removed.size, "", "").length !== removed.size ||
      text.length !== inserted.size
    )
      return null;
    const expected = baseline.tr.insertText(text, start + from, start + oldEnd);
    if (!expected.doc.eq(proposed.doc)) return null;
    return { from: start + from, to: start + oldEnd, text, semantic: "composition" as const };
  };
  const applyNative = (view: EditorView, transaction: Transaction) =>
    Result.try({
      try: () => view.state.applyTransaction(transaction),
      catch: (cause) =>
        new CanonicalSessionError({
          gap: CANONICAL_GAP.dispatch,
          reason: "refused",
          message: `Composition projection failed: ${String(cause)}`,
        }),
    });
  const finish = (view: EditorView, commit: boolean): CompositionFinish => {
    clearTimer();
    authorizedNativeState = null;
    const pending = state;
    if (pending.type === "committed") return { type: "idle" };
    const input =
      commit && pending.type === "provisional" ? replacement(pending.baseline, view.state) : null;
    // A final beforeinput carries the whole selected replacement, even when
    // the native observer produced a smaller diff with a shared prefix/suffix.
    // Acceptance is validated by replacement(); payload identity depends on text,
    // including when the accepted minimal diff retains non-inclusive metadata.
    const committedText = selectedText(pending.baseline, view.state);
    const proposed = view.state;
    const from = pending.baseline.selection.from;
    completed = { type: "idle" };
    state = { type: "committed" };
    if (!view.isDestroyed) {
      // Recovery may finish without a native end event (for example a refused
      // final beforeinput). Let PM end its own lifecycle before restoring DOM.
      if (view.composing) view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
      view.updateState(pending.baseline);
    }
    options.end();
    if (!commit || view.isDestroyed) return { type: "cancelled" };
    if (pending.type === "refused") return { type: "refused", text: pending.text };
    if (!input) {
      options.refuse(
        "Composition changed unsupported content; the canonical document was restored.",
      );
      return { type: "refused", text: committedText };
    }
    if (input.from !== input.to || input.text !== "") options.replace(input);
    if (
      committedText !== null &&
      !view.isDestroyed &&
      view.state.doc.textContent === proposed.doc.textContent &&
      from + committedText.length <= view.state.doc.content.size &&
      view.state.doc.textBetween(from, from + committedText.length, "", "") === committedText
    ) {
      completed = {
        type: "completed",
        receipt: {
          view,
          state: view.state,
          from,
          to: from + committedText.length,
          text: committedText,
          marks:
            pending.baseline.storedMarks ??
            pending.baseline.selection.$from.marksAcross(pending.baseline.selection.$to),
        },
      };
    }
    return { type: "committed", text: committedText };
  };
  const schedule = (view: EditorView) => {
    clearTimer();
    // Allow the native final input and PM's delayed composition flush to settle.
    timer = setCanonicalInputTimer(() => finish(view, true), CANONICAL_COMPOSITION_FLUSH_DELAY_MS);
  };
  return {
    // Keep declaration emission tied to the authored union, not inferred member order.
    get status(): CompositionState["type"] {
      return state.type;
    },
    get active() {
      return state.type !== "committed";
    },
    matchesCancellationRange: (view: EditorView, range: { from: number; to: number }) => {
      if (state.type === "committed" || state.view !== view) return false;
      const text = selectedText(state.baseline, view.state);
      if (text === null || text.length === 0) return false;
      const from = state.baseline.selection.from;
      if (range.from !== from || range.to !== from + text.length) return false;
      // A one-character Backspace can target the whole proposal too. Native
      // cancellation selects that proposal; ordinary deletion leaves a caret.
      const selection = view.dom.ownerDocument.getSelection();
      if (!selection || selection.rangeCount !== 1) return false;
      const selected = selection.getRangeAt(0);
      if (!view.dom.contains(selected.startContainer) || !view.dom.contains(selected.endContainer))
        return false;
      return (
        view.posAtDOM(selected.startContainer, selected.startOffset) === range.from &&
        view.posAtDOM(selected.endContainer, selected.endOffset) === range.to
      );
    },
    start: (view: EditorView) => {
      if (state.type !== "committed") return;
      if (!options.begin()) return;
      state = { type: "provisional", baseline: view.state, view, phase: "active" };
    },
    ended: (view: EditorView) => {
      if (state.type === "committed") return;
      switch (state.type) {
        case "provisional":
          state = {
            type: "provisional",
            baseline: state.baseline,
            view: state.view,
            phase: "ended",
          };
          break;
        case "refused":
          state = {
            type: "refused",
            baseline: state.baseline,
            view: state.view,
            phase: "ended",
            text: state.text,
          };
          break;
        default:
          state satisfies never;
      }
      schedule(view);
    },
    forgetCompleted: () => {
      completed = { type: "idle" };
    },
    get pendingFinal() {
      return completed.type === "authorized";
    },
    authorizeLateFinal: (view: EditorView, text: string | null) => {
      if (completed.type === "idle") return false;
      const { receipt } = completed;
      if (
        view.isDestroyed ||
        receipt.view !== view ||
        receipt.state !== view.state ||
        text === null
      ) {
        completed = { type: "idle" };
        return false;
      }
      completed = { type: "authorized", receipt, text };
      return true;
    },
    authorizeNative: (view: EditorView) => {
      authorizedNativeState = view.state;
    },
    flushed: (view: EditorView) => {
      const authorized = authorizedNativeState;
      // A classified late final remains bound to its state until the observer
      // consumes it or the next gesture invalidates it, including delayed flushes.
      queueMicrotask(() => {
        if (authorizedNativeState === authorized) authorizedNativeState = null;
      });
      if (state.type !== "committed" && state.phase === "ended") schedule(view);
    },
    recover: (view: EditorView) => finish(view, true),
    cancel: (view: EditorView) => finish(view, false),
    reset: () => {
      completed = { type: "idle" };
      if (state.type !== "committed") finish(state.view, false);
      else {
        clearTimer();
        authorizedNativeState = null;
      }
    },
    accept: (view: EditorView, transaction: Transaction) => {
      if (state.type === "committed") {
        const final = completed;
        if (final.type !== "authorized" || final.receipt.view !== view) return false;
        completed = { type: "idle" };
        if (view.isDestroyed || final.receipt.state !== view.state) return false;
        const applied = applyNative(view, transaction);
        const baseline = EditorState.create({
          schema: final.receipt.state.schema,
          doc: final.receipt.state.doc,
          selection: TextSelection.create(
            final.receipt.state.doc,
            final.receipt.from,
            final.receipt.to,
          ),
          storedMarks: final.receipt.marks,
        });
        const input = applied.isOk() ? replacement(baseline, applied.value.state) : null;
        // A final can arrive as the full replacement or its minimal native diff.
        // The captured range and full payload must match in either form.
        if (
          applied.isErr() ||
          !input ||
          input.from < final.receipt.from ||
          input.to > final.receipt.to ||
          selectedText(baseline, applied.value.state) !== final.text
        ) {
          view.updateState(final.receipt.state);
          options.refuse(
            applied.isErr()
              ? applied.error.message
              : "Final composition changed unsupported content; the canonical document was restored.",
          );
          return true;
        }
        if (final.receipt.text !== final.text)
          options.replace({
            from: final.receipt.from,
            to: final.receipt.to,
            text: final.text,
            semantic: "composition",
            compositionPhase: "correction",
          });
        if (!view.isDestroyed) view.updateState(view.state);
        if (
          !view.isDestroyed &&
          view.state.doc.textContent === applied.value.state.doc.textContent &&
          final.receipt.from + final.text.length <= view.state.doc.content.size &&
          view.state.doc.textBetween(
            final.receipt.from,
            final.receipt.from + final.text.length,
            "",
            "",
          ) === final.text
        ) {
          completed = {
            type: "completed",
            receipt: {
              view,
              state: view.state,
              from: final.receipt.from,
              to: final.receipt.from + final.text.length,
              text: final.text,
              marks: final.receipt.marks,
            },
          };
        }
        return true;
      }
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
      const applied = applyNative(view, transaction);
      if (applied.isErr()) {
        view.updateState(state.baseline);
        state = {
          type: "refused",
          baseline: state.baseline,
          view: state.view,
          phase: state.phase,
          text: selectedText(state.baseline, transaction),
        };
        options.refuse(applied.error.message);
        return true;
      }
      // Validate the entire proposal against one plain replacement, including plugin output.
      if (!replacement(state.baseline, applied.value.state)) {
        view.updateState(state.baseline);
        state = {
          type: "refused",
          baseline: state.baseline,
          view: state.view,
          phase: state.phase,
          text: selectedText(state.baseline, applied.value.state),
        };
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
