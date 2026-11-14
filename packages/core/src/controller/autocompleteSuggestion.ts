/**
 * The inline autocomplete suggestion: ghost text a host streams in at a text
 * position, shown beside the caret until the person accepts or dismisses it.
 *
 * The suggestion is view state. It never touches the document, its history
 * or its pending changes; only accepting it edits the document, as typed
 * text. Its anchor is a text position, so it stays on its text whatever the
 * editor or the layout does: an edit a peer makes in a shared room moves it
 * as it moves the person's caret, typing the ghost's next characters at the
 * anchor consumes them and the rest follows the caret, and any other edit
 * of the person's dismisses it.
 */

import type { EditorState, Transaction } from "prosemirror-state";
import { ReplaceStep } from "prosemirror-transform";

import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  sameStory,
  type TextPosition,
  transformPosition,
} from "@stll/docx-core/ops";

import type { CanonicalOrigin } from "../canonical/canonicalSession";
import type { FolioAutocompleteState } from "../types/autocomplete";
import { legacyTextAddressAt } from "./legacyTextAddress";

export type {
  FolioAutocompleteApi,
  FolioAutocompleteState,
} from "../types/autocomplete";

/** No suggestion. */
export const IDLE_AUTOCOMPLETE: FolioAutocompleteState = Object.freeze({
  status: "idle",
});

/** What an edit that changed the document means for the suggestion. */
export type AutocompleteEdit =
  /** The person inserted `text` at `at`, and nothing else. */
  | { type: "typed"; at: TextPosition; text: string }
  /** Another actor's edit, which moves text positions. */
  | { type: "moved"; ops: readonly DocumentOp[] }
  /** Any other edit. */
  | { type: "other" };

export type AutocompleteSuggestion = {
  get: () => FolioAutocompleteState;
  /** Start a suggestion at `anchor`, replacing any other. */
  start: (anchor: TextPosition, requestId: string) => void;
  /** Add streamed text to the suggestion `requestId` names. */
  append: (requestId: string, delta: string) => boolean;
  /** End streaming; a suggestion without text ends with it. */
  finish: (requestId: string) => boolean;
  clear: () => void;
  /** Follow an edit that changed the document. */
  noteEdit: (edit: AutocompleteEdit) => void;
  /** Called after every change of the state; returns the unsubscribe. */
  subscribe: (listener: (state: FolioAutocompleteState) => void) => () => void;
};

/** A canonical commit: the person's single insertion, a peer's edit, or anything else. */
export const classifyAutocompleteEdit = ({
  ops,
  origin,
}: {
  ops: readonly DocumentOp[];
  origin: CanonicalOrigin;
}): AutocompleteEdit => {
  if (origin === "remote") return { type: "moved", ops };
  if (origin !== "input") return { type: "other" };
  const [op, ...rest] = ops;
  return op?.type === DOCUMENT_OP_TYPES.INSERT_TEXT && rest.length === 0
    ? { type: "typed", at: op.at, text: op.text }
    : { type: "other" };
};

/**
 * A default-session edit: plain text inserted at one point, or anything else.
 */
export const legacyAutocompleteEdit = (
  transaction: Transaction,
): AutocompleteEdit => {
  const [step, ...rest] = transaction.steps;
  if (
    !(step instanceof ReplaceStep) ||
    rest.length > 0 ||
    step.from !== step.to ||
    step.slice.openStart !== 0 ||
    step.slice.openEnd !== 0
  )
    return { type: "other" };
  let text = "";
  let plain = true;
  step.slice.content.forEach((node) => {
    if (node.isText) text += node.text ?? "";
    else plain = false;
  });
  const at = legacyTextAddressAt(transaction.before, step.from);
  return plain && at !== null ? { type: "typed", at, text } : { type: "other" };
};

/** The text a word-wise accept takes: up to and including the next word and one space after it. */
export const nextAutocompleteWord = (text: string): string => {
  const match = /^\s*\S+\s?/u.exec(text);
  return match === null ? text : match[0];
};

/** The suggestion owner of one editor. */
export const createAutocompleteSuggestion = (): AutocompleteSuggestion => {
  let state: FolioAutocompleteState = IDLE_AUTOCOMPLETE;
  const listeners = new Set<(next: FolioAutocompleteState) => void>();
  const set = (next: FolioAutocompleteState): void => {
    if (next === state) return;
    state = next;
    for (const listener of listeners) listener(next);
  };
  const active = (requestId: string) =>
    state.status !== "idle" && state.requestId === requestId ? state : null;

  return {
    get: () => state,
    start: (anchor, requestId) =>
      set({ status: "streaming", requestId, anchor, text: "" }),
    append: (requestId, delta) => {
      const current = active(requestId);
      if (current === null) return false;
      if (delta.length > 0)
        set({ ...current, status: "streaming", text: current.text + delta });
      return true;
    },
    finish: (requestId) => {
      const current = active(requestId);
      if (current === null) return false;
      set(
        current.text.length === 0
          ? IDLE_AUTOCOMPLETE
          : { ...current, status: "shown" },
      );
      return true;
    },
    clear: () => set(IDLE_AUTOCOMPLETE),
    noteEdit: (edit) => {
      if (state.status === "idle") return;
      const current = state;
      switch (edit.type) {
        case "typed": {
          const atAnchor =
            sameStory(edit.at.story, current.anchor.story) &&
            edit.at.blockId === current.anchor.blockId &&
            edit.at.offset === current.anchor.offset;
          if (
            !atAnchor ||
            edit.text.length === 0 ||
            !current.text.startsWith(edit.text)
          ) {
            set(IDLE_AUTOCOMPLETE);
            return;
          }
          const rest = current.text.slice(edit.text.length);
          set(
            rest.length === 0 && current.status === "shown"
              ? IDLE_AUTOCOMPLETE
              : {
                  ...current,
                  anchor: {
                    story: current.anchor.story,
                    blockId: current.anchor.blockId,
                    offset: current.anchor.offset + edit.text.length,
                  },
                  text: rest,
                },
          );
          return;
        }
        case "moved": {
          // A peer's text at the anchor's own offset lands after it, as it
          // lands after the person's caret.
          const anchor = transformPosition(
            current.anchor,
            [
              {
                schema: DOCUMENT_OP_SCHEMA_VERSION,
                opId: "autocomplete-anchor",
                actor: "peer",
                baseRev: 0,
                revision: 1,
                ops: edit.ops,
              },
            ],
            "before",
          );
          set(
            anchor === undefined ? IDLE_AUTOCOMPLETE : { ...current, anchor },
          );
          return;
        }
        case "other":
          set(IDLE_AUTOCOMPLETE);
          return;
        default: {
          const unreachable: never = edit;
          return unreachable;
        }
      }
    },
    subscribe: (listener) => {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
  };
};

type AutocompleteTriggerSkipReason =
  "selection-non-empty" | "midword" | "deadzone" | "empty-doc";

type AutocompleteTriggerCheck =
  | { ok: true }
  | { ok: false; reason: AutocompleteTriggerSkipReason; detail?: string };

/** Node type names where no suggestion is offered. */
const AUTOCOMPLETE_DEAD_ZONE_NODES: ReadonlySet<string> = new Set([
  "heading",
  "title",
  "code",
  "code_block",
  "codeBlock",
  "horizontal_rule",
  "horizontalRule",
  "math_block",
  "math_inline",
  "inline_math",
  "inlineMath",
  "equation",
  "citation",
  "citationPreview",
  "pendingCitation",
  "image",
  "imageBlock",
]);

// Letters, numbers, and a small set of in-word connectors. The
// regex uses Unicode property escapes so accented Czech/German/
// Polish letters count as word chars and don't trip the gate.
const WORD_CONTINUATION_CHAR = /[\p{L}\p{N}_'-]/u;

/**
 * Decide whether autocomplete should fire at the current cursor.
 * Designed to be called from a debounced trigger in the host app
 * before it asks for a suggestion. Returns a structured
 * skip reason rather than a bare boolean so callers (and debug
 * UIs) can explain why no suggestion appeared.
 *
 * Gates applied:
 *  - selection must be a collapsed caret (no range selection)
 *  - the cursor must not be inside a dead-zone node (heading,
 *    code, citation, math, etc.) anywhere in the ancestor chain
 *  - the character immediately before the cursor must not be a
 *    word-continuation char — if it is, the model would be
 *    completing a partial token and almost always misreads intent
 */
export const canTriggerAutocomplete = (
  state: EditorState,
): AutocompleteTriggerCheck => {
  const sel = state.selection;
  if (!sel.empty) {
    return { ok: false, reason: "selection-non-empty" };
  }
  if (state.doc.content.size === 0) {
    return { ok: false, reason: "empty-doc" };
  }
  const $cursor = sel.$from;
  for (let depth = $cursor.depth; depth >= 0; depth--) {
    const nodeName = $cursor.node(depth).type.name;
    if (AUTOCOMPLETE_DEAD_ZONE_NODES.has(nodeName)) {
      return { ok: false, reason: "deadzone", detail: nodeName };
    }
  }
  const parentOffset = $cursor.parentOffset;
  if (parentOffset > 0) {
    const before = $cursor.parent.textBetween(
      parentOffset - 1,
      parentOffset,
      "\n",
      "\n",
    );
    if (before.length > 0 && WORD_CONTINUATION_CHAR.test(before)) {
      return { ok: false, reason: "midword", detail: before };
    }
  }
  return { ok: true };
};
