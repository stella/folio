import type {
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
  FolioDocumentOperationBatch,
  FolioDocumentOperationResult,
} from "@stll/folio-core/server";
import {
  assertSupportedFolioDocumentOperationVersion,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
} from "@stll/folio-core/server";
import type { FolioCommentAnchor, FolioReviewChange } from "@stll/folio-core/ai-edits";
import { createReply } from "@stll/folio-core/docx/replyToComment";
import type { Comment } from "@stll/folio-core/types/content";

import type { FolioAgentBridge } from "../bridge";
import type { FolioAgentChange, FolioAgentComment, FolioAgentCommentReply } from "../types";
import { toAgentChange } from "./shared";

export type FolioAgentEditorApplyDocumentOperationsOptions = {
  snapshot: FolioAIEditSnapshot;
  batch: FolioDocumentOperationBatch;
  author?: string;
};

/**
 * Minimal structural slice of `DocxEditorRef` (`packages/react`) this bridge
 * drives — declared locally per AGENTS.md's react-free-core rule (agents
 * stays framework-neutral; it may not import a React-package type), covering
 * ONLY the members actually called below.
 *
 * The read-surface members (`getTrackedChanges`, `getCommentAnchors`,
 * `getSelectionText`, `getPageText`) are OPTIONAL here even though the
 * current `DocxEditorRef` always implements them: a ref built against an
 * older `@stll/folio-react` (before these methods existed) still
 * structurally satisfies this type, and `createEditorRefBridge` below falls
 * back to the pre-existing degraded behavior for each one it does not find.
 */
export type FolioAgentEditorRefLike = {
  /** `DocxEditorRef.createAIEditSnapshot`. `null` before the editor view mounts. */
  createAIEditSnapshot(): FolioAIEditSnapshot | null;
  /** `DocxEditorRef.applyAIEditOperations`. */
  applyAIEditOperations(options: {
    snapshot: FolioAIEditSnapshot;
    operations: FolioAIEditOperation[];
    mode?: FolioAIEditApplyMode;
    author?: string;
  }): FolioAIEditApplyResult;
  /** `DocxEditorRef.applyDocumentOperations`, when available on newer refs. */
  applyDocumentOperations?(
    options: FolioAgentEditorApplyDocumentOperationsOptions,
  ): FolioDocumentOperationResult;
  /** `DocxEditorRef.scrollToBlock`. */
  scrollToBlock(blockId: string, snapshot?: FolioAIEditSnapshot): boolean;
  /** `DocxEditorRef.getTotalPages`. */
  getTotalPages(): number;
  /**
   * `DocxEditorRef.getTrackedChanges`. When present, `getChanges()` maps its
   * output to {@link FolioAgentChange} (same mapping as the reviewer
   * bridge's); when absent, `getChanges()` returns `[]`.
   */
  getTrackedChanges?(): FolioReviewChange[];
  /**
   * `DocxEditorRef.getCommentAnchors`. When present, `getComments()` merges
   * each anchor's `blockId` / `quote` into the matching host-state comment
   * by `commentId`; when absent, those fields stay `null` / `""`.
   */
  getCommentAnchors?(): FolioCommentAnchor[];
  /**
   * `DocxEditorRef.getSelectionText`. The bridge exposes `getSelectionText`
   * only when this is present, so `read_selection` reports an
   * unsupported-capability error on a ref that lacks it.
   */
  getSelectionText?(): string;
  /**
   * `DocxEditorRef.getPageText`. The bridge exposes `getPageText` only when
   * this is present, so `read_page` reports an unsupported-capability error
   * on a ref that lacks it. See {@link createEditorRefBridge} for how a
   * `null` return (page in range, layout not yet computed) is handled.
   */
  getPageText?(page: number): string | null;
};

/** Options for {@link createEditorRefBridge}. */
export type CreateEditorRefBridgeOptions = {
  ref: FolioAgentEditorRefLike;
  /** Author attributed to tracked changes, comments, and replies this bridge creates. */
  author: string;
  /** Read the host app's current comment state (e.g. the `DocxEditor` `comments` prop). */
  getComments(): Comment[];
  /** Replace the host app's comment state (e.g. the setter backing that same prop). */
  setComments(comments: Comment[]): void;
  /** `"tracked-changes"` (default) produces ins/del redlines; `"direct"` edits in place. */
  mode?: FolioAIEditApplyMode;
};

const toAgentCommentReply = (reply: Comment): FolioAgentCommentReply => ({
  id: String(reply.id),
  author: reply.author,
  text: replyPlainText(reply),
});

/**
 * Parse a tool-supplied comment id into the numeric `Comment.id` this bridge
 * matches against, rejecting anything but a bare non-negative integer.
 * `Number.parseInt` alone would silently accept trailing junk (`"12abc"` ->
 * `12`), which could reply to or resolve the wrong comment on malformed tool
 * input.
 */
const parseCommentId = (commentId: string): number | null =>
  /^\d+$/.test(commentId) ? Number.parseInt(commentId, 10) : null;

/** A host-state `Comment`'s plain text, its paragraphs joined by newlines (mirrors `FolioDocxReviewer`'s reading). */
const replyPlainText = (comment: Comment): string =>
  comment.content.map(paragraphPlainText).join("\n");

const paragraphPlainText = (paragraph: Comment["content"][number]): string => {
  const parts: string[] = [];
  for (const item of paragraph.content ?? []) {
    if (item.type !== "run") {
      continue;
    }
    for (const runItem of item.content ?? []) {
      if (runItem.type === "text") {
        parts.push(runItem.text);
      }
    }
  }
  return parts.join("");
};

/**
 * Build a {@link FolioAgentBridge} over a live `DocxEditor` ref plus the host
 * app's comment state.
 *
 * Comments live in app-controlled React state (the `DocxEditor` `comments`
 * prop), not on the ref, so this factory takes `getComments`/`setComments` to
 * read and write that state directly — the same pair the host already passes
 * to `DocxEditor`.
 *
 * KNOWN LIMITATIONS (only apply to a `ref` that predates the read-surface
 * additions below; the current `DocxEditorRef` implements all four):
 * - `getChanges()` returns `[]` when `ref.getTrackedChanges` is absent, since
 *   there is then no ref-level way to enumerate tracked changes from
 *   ProseMirror mark attributes.
 * - Comment entries fall back to `blockId: null` / `quote: ""` when
 *   `ref.getCommentAnchors` is absent, since there is then no ref-level way
 *   to resolve a comment's anchor against the live ProseMirror document.
 * - `read_page` / `read_selection` report an unsupported-capability error
 *   when `ref.getPageText` / `ref.getSelectionText` are absent: this bridge
 *   omits the corresponding `getPageText` / `getSelectionText` member
 *   entirely rather than implementing it as a no-op, which is what tells
 *   `executeFolioToolCall` to report the tool as unsupported instead of
 *   throwing.
 */
export const createEditorRefBridge = (options: CreateEditorRefBridgeOptions): FolioAgentBridge => {
  const { ref, author, getComments, setComments } = options;
  const mode = options.mode ?? "tracked-changes";

  const requireSnapshot = (): FolioAIEditSnapshot => {
    const snapshot = ref.createAIEditSnapshot();
    if (!snapshot) {
      throw new Error("The editor view is not mounted; no snapshot is available yet.");
    }
    return snapshot;
  };

  const bridge: FolioAgentBridge = {
    snapshot: requireSnapshot,
    applyDocumentOperations: (batch) => {
      assertSupportedFolioDocumentOperationVersion(batch.version);
      const snapshot = requireSnapshot();
      const versionedBatch = { ...batch, mode };
      if (ref.applyDocumentOperations) {
        return ref.applyDocumentOperations({ snapshot, batch: versionedBatch, author });
      }
      if (versionedBatch.atomic === true) {
        return {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          status: "rejected",
          applied: [],
          skipped: versionedBatch.operations.map(({ id }) => ({
            id,
            reason: "unsupportedMode",
          })),
        };
      }
      return {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        status: "committed",
        ...ref.applyAIEditOperations({
          snapshot,
          operations: versionedBatch.operations,
          mode: versionedBatch.mode,
          author,
        }),
      };
    },
    getComments: (): FolioAgentComment[] => {
      const comments = getComments();
      const anchors = ref.getCommentAnchors?.();
      const anchorByCommentId = new Map<number, FolioCommentAnchor>();
      if (anchors) {
        for (const anchor of anchors) {
          anchorByCommentId.set(anchor.commentId, anchor);
        }
      }
      const repliesByParent = new Map<number, Comment[]>();
      const topLevel: Comment[] = [];
      for (const comment of comments) {
        // `parentId` is typed `number | undefined`, but hosts that round-trip
        // comments through JSON (or Word-imported state) can hand back
        // `parentId: null` for a top-level comment; treat both as "no parent"
        // (see `getCommentParentId` in packages/react/commentsHelpers.ts).
        if (comment.parentId == null) {
          topLevel.push(comment);
          continue;
        }
        const siblings = repliesByParent.get(comment.parentId) ?? [];
        siblings.push(comment);
        repliesByParent.set(comment.parentId, siblings);
      }
      return topLevel.map((comment) => {
        const anchor = anchorByCommentId.get(comment.id);
        return {
          id: String(comment.id),
          author: comment.author,
          text: replyPlainText(comment),
          resolved: comment.done ?? false,
          // Merged from `ref.getCommentAnchors()` when the ref provides it;
          // `null` / `""` on an older ref — see KNOWN LIMITATIONS above.
          blockId: anchor?.blockId ?? null,
          quote: anchor?.quote ?? "",
          replies: (repliesByParent.get(comment.id) ?? []).map(toAgentCommentReply),
        };
      });
    },
    getChanges: (): FolioAgentChange[] => {
      const changes = ref.getTrackedChanges?.();
      return changes ? changes.map(toAgentChange) : [];
    },
    replyToComment: (commentId, text) => {
      const parentId = parseCommentId(commentId);
      if (parentId === null) {
        return false;
      }
      const comments = getComments();
      const reply = createReply(comments, parentId, { author, text });
      if (!reply) {
        return false;
      }
      setComments([...comments, reply]);
      return true;
    },
    resolveComment: (commentId, resolved) => {
      const targetId = parseCommentId(commentId);
      if (targetId === null) {
        return false;
      }
      const comments = getComments();
      let found = false;
      const next = comments.map((comment) => {
        if (comment.id !== targetId) {
          return comment;
        }
        found = true;
        return { ...comment, done: resolved };
      });
      if (!found) {
        return false;
      }
      setComments(next);
      return true;
    },
    scrollToBlock: (blockId) => ref.scrollToBlock(blockId),
    getPageCount: () => ref.getTotalPages(),
  };

  const getSelectionText = ref.getSelectionText;
  if (getSelectionText) {
    bridge.getSelectionText = () => getSelectionText();
  }

  const getPageText = ref.getPageText;
  if (getPageText) {
    // `ref.getPageText` returns `null` when the page number is in range but
    // its layout has not been computed yet (a transient render-timing gap,
    // not an error) — `read_page` in execute.ts already rejects a page
    // number beyond `getPageCount()` before ever calling this, so a `null`
    // here means "not rendered yet", not "out of range". Map it to `""`: an
    // empty page string is a safe, retryable degrade for a model mid tool
    // call, instead of throwing out of a call it has no way to recover from.
    bridge.getPageText = (page) => getPageText(page) ?? "";
  }

  return bridge;
};
