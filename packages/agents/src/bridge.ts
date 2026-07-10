import type {
  FolioAIEditSnapshot,
  FolioDocumentOperationBatch,
  FolioDocumentOperationResult,
  FolioDocumentStory,
  FolioDocumentStoryHandle,
} from "@stll/folio-core/server";

import type { FolioAgentChange, FolioAgentComment } from "./types";

/**
 * The structural contract every folio surface (headless reviewer, live
 * editor) implements so the tool executor can drive either one identically.
 * `bridges/reviewer.ts` and `bridges/editor-ref.ts` are the two shipped
 * implementations; a host app can also hand-write one.
 *
 * The required members are available on every surface. The optional members
 * are live-editor-only capabilities (paging, selection, scrolling): when a
 * bridge omits one, {@link executeFolioToolCall} reports the corresponding
 * tool call as an unsupported-capability error rather than throwing, so a
 * model driving a headless reviewer gets a plain-language reason instead of a
 * crash.
 */
export type FolioAgentBridge = {
  /** Snapshot the current document into AI-facing blocks + anchors. */
  snapshot(): FolioAIEditSnapshot;
  /**
   * Apply a versioned operation batch against the current document. The
   * bridge decides mode (tracked-changes by default) and author internally.
   */
  applyDocumentOperations(batch: FolioDocumentOperationBatch): FolioDocumentOperationResult;
  /** The comment threads present in the document. */
  getComments(): FolioAgentComment[];
  /** The pending tracked changes (insertions/deletions) present in the document. */
  getChanges(): FolioAgentChange[];
  /** Discover typed document stories when the surface exposes package parts. */
  listStories?(): FolioDocumentStory[];
  /** Read one previously discovered story. */
  readStory?(handle: FolioDocumentStoryHandle): FolioDocumentStory | null;
  /** Reply to a comment thread. Returns `false` when the target comment does not exist. */
  replyToComment(commentId: string, text: string): boolean;
  /** Mark a comment thread resolved or reopen it. Returns `false` when the target comment does not exist. */
  resolveComment(commentId: string, resolved: boolean): boolean;

  // ---------------------------------------------------------------------
  // Optional capabilities — live editor only. Omit a member entirely (do
  // not implement it as a no-op) to signal the executor should treat the
  // corresponding tool as unsupported on this surface.
  // ---------------------------------------------------------------------

  /** Scroll the live editor to the given block and select it. */
  scrollToBlock?(blockId: string): boolean;
  /** The user's current text selection in the live editor, as plain text. */
  getSelectionText?(): string;
  /** Total page count in the live, paginated editor. */
  getPageCount?(): number;
  /** Plain text of the given 1-based page in the live editor. */
  getPageText?(page: number): string;
};
