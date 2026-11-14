import type { TextPosition } from "@stll/docx-core/ops";

/**
 * The inline autocomplete suggestion: none, or ghost text at a text position
 * that is still streaming in or complete.
 */
export type FolioAutocompleteState =
  | { status: "idle" }
  | {
      status: "streaming" | "shown";
      /** Names the suggestion in `append` and `finish`. */
      requestId: string;
      /** Where the ghost text would be inserted. */
      anchor: TextPosition;
      /** The ghost text not yet typed or accepted. */
      text: string;
    };

/** Drives the inline autocomplete suggestion of one editor. */
export type FolioAutocompleteApi = {
  /**
   * Start a suggestion at `at` (the caret by default), replacing any other.
   * Returns its request id (`requestId`, or a fresh one), or null when there
   * is no editor or `at` names no paragraph.
   */
  start: (options?: { at?: TextPosition; requestId?: string }) => string | null;
  /** Add streamed text; false when `requestId` is not the current suggestion. */
  append: (requestId: string, delta: string) => boolean;
  /** End streaming; a suggestion without text is dismissed. */
  finish: (requestId: string) => boolean;
  /** Dismiss the suggestion. */
  clear: () => void;
  /**
   * Insert the suggestion, or its next word, as typed text at its anchor:
   * one undo step, a tracked insertion when suggesting. The rest of a
   * word-wise accept stays shown. False when nothing was inserted.
   */
  accept: (scope?: "all" | "word") => boolean;
  get: () => FolioAutocompleteState;
  /**
   * Whether the caret is where a suggestion fits: a collapsed selection, not
   * inside a word, not in a heading, title, code or other excluded block.
   */
  canTrigger: () => boolean;
};
