import { Result } from "better-result";
import {
  DOMParser as ProseMirrorDOMParser,
  DOMSerializer,
  Fragment,
  type Mark,
  type Node as PMNode,
} from "prosemirror-model";
import {
  EditorState,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { splitsSurrogatePair } from "../ai-edits/character-boundaries";
import { CanonicalSessionError } from "../canonical/canonicalSession";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import {
  CANONICAL_COMPOSITION_FLUSH_DELAY_MS,
  setCanonicalInputTimer,
} from "./canonicalInputTimer";

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
  /**
   * The view shows provisional composition text, or the baseline again, that
   * no commit published: it needs a layout pass but changes no document.
   */
  shown?: () => void;
};

type CompositionCompletionEvidence = "provisional" | "explicitFinal";
type CompositionPhase =
  | { type: "active"; completion: CompositionCompletionEvidence }
  | {
      type: "ended";
      data: string | null;
      completion: CompositionCompletionEvidence;
    };

type CompositionState =
  | { type: "committed" }
  | {
      type: "provisional";
      baseline: EditorState;
      view: EditorView;
      phase: CompositionPhase;
    }
  | {
      type: "refused";
      baseline: EditorState;
      view: EditorView;
      phase: CompositionPhase;
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

const REVIEW_MARK_NAMES: ReadonlySet<string> = new Set([
  "insertion",
  "deletion",
]);
const readBackCache = new WeakMap<Mark, Mark | null>();
/**
 * The mark a DOM read rebuilds from the mark's own rendering: `null` when no
 * rule reads it back, `undefined` with no document to render into. A read
 * keeps only what the rendering carries, so this is all a reread can say.
 */
const readBack = (mark: Mark): Mark | null | undefined => {
  const cached = readBackCache.get(mark);
  if (cached !== undefined) return cached;
  if (typeof document === "undefined") return undefined;
  const { schema } = mark.type;
  const container = document.createElement("div");
  container.append(
    DOMSerializer.fromSchema(schema).serializeFragment(
      Fragment.from(schema.text("x", [mark])),
      { document },
    ),
  );
  const texts: PMNode[] = [];
  ProseMirrorDOMParser.fromSchema(schema)
    .parseSlice(container)
    .content.descendants((node) => {
      if (node.isText) texts.push(node);
      return texts.length === 0;
    });
  const read = mark.type.isInSet(texts.at(0)?.marks ?? []) ?? null;
  readBackCache.set(mark, read);
  return read;
};

type NativeReplacementComparisonOptions = {
  baseline: EditorState;
  proposed: EditorState;
  from: number;
  to: number;
  text: string;
};
/**
 * Whether a native proposal is one plain replacement. The commit is a text
 * replacement over the captured model, so this only has to tell DOM noise from a
 * real change. Reading the DOM back can lose a mark no rule reads back or, through
 * `insertText`, a non-inclusive one. It rebuilds any mark from what its rendering
 * carries, a review mark from fewer attributes, and composed text sits in the DOM
 * wrapper of the tracked run beside it. Gaining any other mark, losing or changing
 * one the rendering carries, or another revision is a real change.
 * So is a selected paragraph boundary that does not disappear exactly.
 */
const matchesNativeReplacement = ({
  baseline,
  proposed,
  from,
  to,
  text,
}: NativeReplacementComparisonOptions) => {
  const expected = baseline.tr.insertText(text, from, to);
  if (expected.doc.eq(proposed.doc)) return true;
  const $from = baseline.doc.resolve(from);
  const $to = baseline.doc.resolve(to);
  if (expected.doc.content.size !== proposed.doc.content.size) return false;
  const position = expected.doc.resolve(from);
  const paragraph = position.parent;
  const start = position.start();
  const native = proposed.doc.nodeAt(start - 1);
  if (native === null || native.type !== paragraph.type) return false;
  // Across paragraphs the browser keeps the last one's element, and the commit
  // keeps the first one's identity.
  if (!$from.sameParent($to) && native.sameMarkup($to.parent))
    expected.setNodeMarkup(start - 1, undefined, native.attrs, native.marks);
  const composedFrom = from;
  const composedTo = from + text.length;
  const types = Object.values(baseline.schema.marks);
  const beside = [$from.nodeBefore, $to.nodeAfter];
  const corrections: (() => void)[] = [];
  let index = 0;
  let offset = 0;
  native.forEach((node, nodeOffset) => {
    const end = nodeOffset + node.nodeSize;
    let cursor = nodeOffset;
    while (cursor < end && index < paragraph.childCount) {
      const authored = paragraph.child(index);
      const authoredEnd = offset + authored.nodeSize;
      const overlapEnd = Math.min(end, authoredEnd);
      if (node.isText && authored.isText) {
        const rangeFrom = start + cursor;
        const rangeTo = start + overlapEnd;
        for (const type of types) {
          const held = type.isInSet(authored.marks);
          const read = type.isInSet(node.marks);
          if (held && !read) {
            const rebuilt = readBack(held);
            if (
              type.spec.inclusive === false ||
              rebuilt === null ||
              (rebuilt === undefined && (type.spec.parseDOM ?? []).length === 0)
            )
              corrections.push(() =>
                expected.removeMark(rangeFrom, rangeTo, type),
              );
            continue;
          }
          // A reread rebuilds a mark from what its rendering carries, so
          // attributes the rendering cannot hold come back as defaults.
          if (
            held &&
            read &&
            !held.eq(read) &&
            !REVIEW_MARK_NAMES.has(type.name)
          ) {
            if (readBack(held)?.eq(read))
              corrections.push(() =>
                expected.addMark(rangeFrom, rangeTo, read),
              );
            continue;
          }
          if (!read || !REVIEW_MARK_NAMES.has(type.name)) continue;
          const sameRevision =
            held?.attrs["revisionId"] === read.attrs["revisionId"];
          const inWrapperBeside =
            rangeFrom >= composedFrom &&
            rangeTo <= composedTo &&
            beside.some((neighbour) => type.isInSet(neighbour?.marks ?? []));
          if (held ? sameRevision : inWrapperBeside)
            corrections.push(() => expected.addMark(rangeFrom, rangeTo, read));
        }
      }
      cursor = overlapEnd;
      if (cursor === authoredEnd) {
        offset = authoredEnd;
        index += 1;
      }
    }
  });
  for (const correction of corrections) correction();
  return expected.doc.eq(proposed.doc);
};

/** Native IME owns the provisional view; the captured canonical projection stays unchanged. */
export const createCanonicalComposition = (
  options: CompositionOptions,
): {
  readonly status: CompositionState["type"];
  readonly active: boolean;
  matchesCancellationRange: (
    view: EditorView,
    range: { from: number; to: number },
  ) => boolean;
  start: (view: EditorView) => void;
  ended: (view: EditorView, data: string | null) => void;
  forgetCompleted: () => void;
  readonly pendingFinal: boolean;
  authorizeLateFinal: (view: EditorView, text: string | null) => boolean;
  authorizeNative: (
    view: EditorView,
    completion: CompositionCompletionEvidence,
    finalText?: string | null,
  ) => void;
  flushed: (view: EditorView) => void;
  recover: (view: EditorView) => CompositionFinish;
  cancel: (view: EditorView) => CompositionFinish;
  reset: () => void;
  accept: (view: EditorView, transaction: Transaction) => boolean;
} => {
  let state: CompositionState = { type: "committed" };
  let authorizedNativeState: EditorState | null = null;
  let completed: CompletedComposition = { type: "idle" };
  let timer: ReturnType<typeof setTimeout> | undefined;
  /** The browser began again mid-composition; its proposals no longer say where. */
  let restarted = false;
  let acceptedUpdate = false;
  let explicitFinal: string | null = null;
  const clearTimer = () => {
    if (timer !== undefined) clearTimeout(timer);
    timer = undefined;
  };
  const show = (view: EditorView, shown: EditorState) => {
    view.updateState(shown);
    options.shown?.();
  };
  const selectedText = (
    baseline: EditorState,
    proposed: Pick<EditorState, "doc">,
  ) => {
    const { from, to } = baseline.selection;
    const size =
      proposed.doc.content.size - baseline.doc.content.size + to - from;
    if (size < 0 || from + size > proposed.doc.content.size) return null;
    const text = proposed.doc.textBetween(from, from + size, "", "");
    return text.length === size ? text : null;
  };
  const selectedReplacement = (
    baseline: EditorState,
    proposed: EditorState,
  ) => {
    const { from, to } = baseline.selection;
    const text = selectedText(baseline, proposed);
    if (
      text === null ||
      !matchesNativeReplacement({ baseline, proposed, from, to, text })
    )
      return null;
    return { from, to, text, semantic: "composition" as const };
  };
  const replacement = (baseline: EditorState, proposed: EditorState) => {
    const { $from, $to } = baseline.selection;
    if (!$from.parent.isTextblock || !$to.parent.isTextblock) return null;
    // Native updates may omit source identity across the whole text carrier,
    // including its untouched suffix. The captured model owns the committed identities.
    const selected = selectedReplacement(baseline, proposed);
    if (selected) {
      if (
        $from.sameParent($to) &&
        selected.text ===
          baseline.doc.textBetween(selected.from, selected.to, "", "")
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
    // identity). Narrow to the text change; reconstruction still checks every
    // semantic mark instead of committing a same-text replacement.
    while (from < oldEnd && from < newEnd && before[from] === after[from])
      from++;
    while (
      oldEnd > from &&
      newEnd > from &&
      before[oldEnd - 1] === after[newEnd - 1]
    ) {
      oldEnd--;
      newEnd--;
    }
    while (
      splitsSurrogatePair(before, from) ||
      splitsSurrogatePair(after, from)
    )
      from--;
    while (
      splitsSurrogatePair(before, oldEnd) ||
      splitsSurrogatePair(after, newEnd)
    ) {
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
    if (
      !matchesNativeReplacement({
        baseline,
        proposed,
        from: start + from,
        to: start + oldEnd,
        text,
      })
    )
      return null;
    return {
      from: start + from,
      to: start + oldEnd,
      text,
      semantic: "composition" as const,
    };
  };
  /**
   * The browser can lose its composition range when the projection redraws a
   * neighbouring node, and then starts again from a position it computes itself.
   * Its final text still names what was composed, and the captured selection says
   * where. A proposal that reaches the caret and holds that text is such a restart
   * even when the restart event went unseen.
   */
  const finalInput = (
    pending: Extract<CompositionState, { type: "provisional" }>,
    native: NonNullable<ReturnType<typeof replacement>>,
  ) => {
    const { phase, baseline } = pending;
    const final = explicitFinal ?? (phase.type === "ended" ? phase.data : null);
    if (!final || native.text === final) return native;
    const { from, to, empty } = baseline.selection;
    const atCaret =
      empty &&
      native.from <= from &&
      native.to >= from &&
      native.text.includes(final);
    if (!restarted && !atCaret) return native;
    return { from, to, text: final, semantic: "composition" as const };
  };
  const applyNative = (view: EditorView, transaction: Transaction) =>
    Result.try({
      try: () => view.state.applyTransaction(transaction),
      catch: (cause) =>
        new CanonicalSessionError({
          ground: { gap: CANONICAL_GAP.dispatch },
          reason: "refused",
          message: `Composition projection failed: ${String(cause)}`,
        }),
    });
  const finish = (view: EditorView, commit: boolean): CompositionFinish => {
    clearTimer();
    authorizedNativeState = null;
    const pending = state;
    if (pending.type === "committed") return { type: "idle" };
    const native =
      commit && pending.type === "provisional"
        ? replacement(pending.baseline, view.state)
        : null;
    const input =
      native && pending.type === "provisional"
        ? finalInput(pending, native)
        : native;
    // A final beforeinput carries the whole selected replacement, even when
    // the native observer produced a smaller diff with a shared prefix/suffix.
    // Acceptance is validated by replacement(); payload identity depends on text,
    // including when the accepted minimal diff retains non-inclusive metadata.
    const committedText =
      input && input !== native
        ? input.text
        : selectedText(pending.baseline, view.state);
    const proposed = view.state;
    const from = pending.baseline.selection.from;
    completed = { type: "idle" };
    state = { type: "committed" };
    if (!view.isDestroyed) {
      // Recovery may finish without a native end event (for example a refused
      // final beforeinput). Let PM end its own lifecycle before restoring DOM.
      if (view.composing)
        view.dom.dispatchEvent(new Event("compositionend", { bubbles: true }));
      view.updateState(pending.baseline);
    }
    options.end();
    if (
      !commit ||
      view.isDestroyed ||
      (pending.phase.type === "ended" &&
        pending.phase.data === "" &&
        pending.phase.completion === "provisional")
    )
      return { type: "cancelled" };
    if (pending.type === "refused")
      return { type: "refused", text: pending.text };
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
      view.state.doc.textBetween(from, from + committedText.length, "", "") ===
        committedText
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
            pending.baseline.selection.$from.marksAcross(
              pending.baseline.selection.$to,
            ),
        },
      };
    }
    return { type: "committed", text: committedText };
  };
  const schedule = (view: EditorView) => {
    clearTimer();
    // Allow the native final input and PM's delayed composition flush to settle.
    timer = setCanonicalInputTimer(
      () => finish(view, true),
      CANONICAL_COMPOSITION_FLUSH_DELAY_MS,
    );
  };
  return {
    // Keep declaration emission tied to the authored union, not inferred member order.
    get status(): CompositionState["type"] {
      return state.type;
    },
    get active() {
      return state.type !== "committed";
    },
    matchesCancellationRange: (
      view: EditorView,
      range: { from: number; to: number },
    ) => {
      if (state.type === "committed" || state.view !== view) return false;
      // A refused composition already restored the document, so the browser
      // removing its proposal must not land on the restored selection.
      if (state.type === "refused") return true;
      const text = selectedText(state.baseline, view.state);
      if (text === null || text.length === 0) return false;
      const from = state.baseline.selection.from;
      if (range.from !== from || range.to !== from + text.length) return false;
      // A one-character Backspace can target the whole proposal too. Native
      // cancellation selects that proposal; ordinary deletion leaves a caret.
      const selection = view.dom.ownerDocument.getSelection();
      if (!selection || selection.rangeCount !== 1) return false;
      const selected = selection.getRangeAt(0);
      if (
        !view.dom.contains(selected.startContainer) ||
        !view.dom.contains(selected.endContainer)
      )
        return false;
      return (
        view.posAtDOM(selected.startContainer, selected.startOffset) ===
          range.from &&
        view.posAtDOM(selected.endContainer, selected.endOffset) === range.to
      );
    },
    start: (view: EditorView) => {
      if (
        state.type === "provisional" &&
        state.phase.type === "active" &&
        acceptedUpdate
      )
        restarted = true;
      if (state.type !== "committed") return;
      if (!options.begin()) return;
      restarted = false;
      acceptedUpdate = false;
      explicitFinal = null;
      // A browser removes a selection as soon as a composition starts over it,
      // before its first update. replacement() still validates the result.
      if (!view.state.selection.empty) authorizedNativeState = view.state;
      state = {
        type: "provisional",
        baseline: view.state,
        view,
        phase: { type: "active", completion: "provisional" },
      };
    },
    ended: (view: EditorView, data: string | null) => {
      if (state.type === "committed") return;
      switch (state.type) {
        case "provisional":
          state = {
            type: "provisional",
            baseline: state.baseline,
            view: state.view,
            phase: { type: "ended", data, completion: state.phase.completion },
          };
          break;
        case "refused":
          state = {
            type: "refused",
            baseline: state.baseline,
            view: state.view,
            phase: { type: "ended", data, completion: state.phase.completion },
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
    authorizeNative: (
      view: EditorView,
      completion: CompositionCompletionEvidence,
      finalText?: string | null,
    ) => {
      if (completion === "explicitFinal" && typeof finalText === "string")
        explicitFinal = finalText;
      authorizedNativeState = view.state;
      if (state.type !== "committed" && completion === "explicitFinal") {
        state.phase =
          state.phase.type === "active"
            ? { type: "active", completion }
            : { type: "ended", data: state.phase.data, completion };
      }
    },
    flushed: (view: EditorView) => {
      const authorized = authorizedNativeState;
      // A classified late final remains bound to its state until the observer
      // consumes it or the next gesture invalidates it, including delayed flushes.
      queueMicrotask(() => {
        if (authorizedNativeState === authorized) authorizedNativeState = null;
      });
      if (state.type !== "committed" && state.phase.type === "ended")
        schedule(view);
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
        if (final.type !== "authorized" || final.receipt.view !== view)
          return false;
        completed = { type: "idle" };
        if (view.isDestroyed || final.receipt.state !== view.state)
          return false;
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
        const input = applied.isOk()
          ? replacement(baseline, applied.value.state)
          : null;
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
          final.receipt.from + final.text.length <=
            view.state.doc.content.size &&
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
        show(view, state.baseline);
        return true;
      }
      if (restarted) {
        show(view, state.baseline);
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
        show(view, state.baseline);
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
        show(view, state.baseline);
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
      acceptedUpdate = true;
      show(view, applied.value.state);
      return true;
    },
  };
};
