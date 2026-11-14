/** The private playground bridge contract shared by both adapter hosts. */
import type {
  FolioDocumentOperationBatch,
  FolioDocumentOperationResult,
} from "@stll/folio-core/server";

export type FolioParityBridge = {
  /** Drive generated batches through the public ref and return the saved package. */
  runGeneratedFlow: (
    source: number[],
    batches: FolioDocumentOperationBatch[],
  ) => Promise<{
    bytes: number[];
    results: FolioDocumentOperationResult[];
  }>;
  /** Total laid-out pages (0 before the first layout). */
  getTotalPages: () => number;
  /** Public page navigation, exercised against the real scrolling viewport. */
  scrollToPage: (pageNumber: number, handle: "document" | "paged") => boolean;
  readScrollViewport: (pageNumber: number) => {
    scrollTop: number;
    clientHeight: number;
    scrollHeight: number;
    rootMatches: boolean;
    pageTop: number;
    pageBottom: number;
    viewportTop: number;
    viewportBottom: number;
  } | null;
  /** Force-create the deferred editor view (no focus steal). */
  ensureView: () => void;
  /** Whether the live ProseMirror view exists yet. */
  hasView: () => boolean;
  /** Concatenated document text (block separators collapse to empty). */
  getDocumentText: () => string;
  /** Text watermark content, or null when none/a non-text watermark is active. */
  getTextWatermark: () => string | null;
  /** Bold the first word of the document. Returns whether the mark applied. */
  boldFirstWord: () => boolean;
  /** Select the first word through the shared editor controller. */
  selectFirstWord: () => boolean;
  /** Count painted range-selection rectangles. */
  countSelectionRects: () => number;
  /** How many times the host's clipboard callback of this kind has fired. */
  getClipboardCallbackCount: (kind: "copy" | "cut" | "paste") => number;
  /** Table properties at the live selection, or null outside a table. */
  getCurrentTableProperties: () => {
    width: number | null;
    widthType: string | null;
    justification: string | null;
  } | null;
  /** Insert a rows×cols table at the selection (core helper). Returns success. */
  insertTable: (rows: number, cols: number) => boolean;
  /** Count `table` nodes in the live document (0 with no live view). */
  countTables: () => number;
  /** Comment the first block with text through the document-operation API. Returns success. */
  commentFirstBlock: () => boolean;
  /** Count painted `[data-comment-id]` anchors in the pages (shared painter attr). */
  countCommentAnchors: () => number;
  /** Block count of the AI-edit snapshot over the live doc (0 with no live view). */
  aiSnapshotBlockCount: () => number;
  pendingSuggestionPersistence: () => {
    exported: number;
    restaged: number;
    stale: number;
    active: number;
    version: number | null;
  };
  /** Painted geometry exposed through the public ref contract. */
  readBlockGeometry: () => {
    rects: {
      snapshotBlockId: string;
      blockId: string;
      page: number;
      top: number;
      height: number;
    }[];
    missingIsNull: boolean;
    hasScrollRoot: boolean;
  };
  /** Reveal the first stable snapshot block and report target/current pages. */
  navigateToFirstBlock: () => {
    shown: boolean;
    targetPage: number;
    currentPage: number;
  };
  /** Plain text of the current live editor selection. */
  getSelectedText: () => string;
  /** Apply and undo one direct document-operation batch; true only when content restores. */
  applyAndUndoDocumentOperation: () => boolean;
  /** Push an anonymization term matching the first word. Returns whether one was pushed. */
  anonymizeFirstWord: () => boolean;
  /** Count painted anonymization highlight rects in the overlay. */
  countAnonymizationRects: () => number;
  /** Start a streaming autocomplete suggestion at the current selection. */
  startAutocomplete: (text: string) => boolean;
  /** Mark the active autocomplete suggestion as complete. */
  finishAutocomplete: () => boolean;
  /** Dismiss the active autocomplete suggestion. */
  clearAutocomplete: () => boolean;
  /** Serialize to DOCX and return the byte length (0 on failure). */
  save: () => Promise<number>;
  /** Open a DOCX package through the public ref, with the editor view ready. */
  loadDocx: (bytes: number[]) => Promise<void>;
  /** Serialize to DOCX through the public ref and return the package. */
  saveDocx: () => Promise<number[]>;
  /**
   * Suggest a paragraph after the first block holding `after`, with a comment
   * on it; returns the suggestion id, or null.
   */
  suggestCommentedInsertion: (
    after: string,
    text: string,
    comment: string,
  ) => string | null;
  /** Reject one suggestion through the public ref. */
  rejectSuggestion: (suggestionId: string) => boolean;
  /** Whether the live editor has edits not yet serialized by save(). */
  hasPendingChanges: () => boolean;
  /**
   * Insert text through `getEditorRef().dispatch` (the nested `PagedEditorRef`
   * handle), not the raw ProseMirror `view.dispatch` the other insert methods
   * use. Exercises the Vue-synthesized ref's `dispatch` method end-to-end.
   */
  insertTextViaPagedEditorRef: (text: string) => boolean;
  /**
   * Page number (1-indexed) containing the current selection anchor, resolved
   * through `getEditorRef().getPageNumberForPmPos`. 0 with no live view/layout.
   */
  getPageNumberForSelection: () => number;
  /**
   * Type `marker` at the selection, then in the same task load the document as
   * it was before the keystroke through `loadDocument`, as a host applying a
   * new revision while the user types does. False without a live view.
   */
  typeThenReloadDocument: (marker: string) => boolean;
};
