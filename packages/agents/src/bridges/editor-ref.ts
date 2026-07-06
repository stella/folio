import type {
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
} from "@stll/folio-core/server";
import { createReply } from "@stll/folio-core/docx/replyToComment";
import type { Comment } from "@stll/folio-core/types/content";

import type { FolioAgentBridge } from "../bridge";
import type { FolioAgentChange, FolioAgentComment, FolioAgentCommentReply } from "../types";

/**
 * Minimal structural slice of `DocxEditorRef` (`packages/react`) this bridge
 * drives — declared locally per AGENTS.md's react-free-core rule (agents
 * stays framework-neutral; it may not import a React-package type), covering
 * ONLY the members actually called below.
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
  /** `DocxEditorRef.scrollToBlock`. */
  scrollToBlock(blockId: string, snapshot?: FolioAIEditSnapshot): boolean;
  /** `DocxEditorRef.getTotalPages`. */
  getTotalPages(): number;
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
 * KNOWN LIMITATIONS (structural gaps in `DocxEditorRef`'s current surface,
 * not bugs here):
 * - `getChanges()` always returns `[]`. Tracked changes are read from
 *   ProseMirror mark attributes (`FolioDocxReviewer.getChanges` walks
 *   `state.doc`), which `DocxEditorRef` does not expose; there is no ref
 *   method to enumerate them from outside the editor.
 * - Comment entries have no `blockId` or `quote`: those come from resolving
 *   each comment's anchor against the live ProseMirror document, which
 *   likewise has no ref-level accessor. Both are set to `null` / `""`.
 * - `read_page` and `read_selection` are unsupported: `DocxEditorRef` has no
 *   `getPageText` / `getSelectionText` equivalent (only `scrollToBlock` and
 *   `getTotalPages`), so those bridge members are omitted entirely and the
 *   corresponding tools report an unsupported-capability error.
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

  return {
    snapshot: requireSnapshot,
    applyOperations: (operations) =>
      ref.applyAIEditOperations({ snapshot: requireSnapshot(), operations, mode, author }),
    getComments: (): FolioAgentComment[] => {
      const comments = getComments();
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
      return topLevel.map((comment) => ({
        id: String(comment.id),
        author: comment.author,
        text: replyPlainText(comment),
        resolved: comment.done ?? false,
        // Not resolvable from host state alone — see KNOWN LIMITATIONS above.
        blockId: null,
        quote: "",
        replies: (repliesByParent.get(comment.id) ?? []).map(toAgentCommentReply),
      }));
    },
    // Not resolvable from `DocxEditorRef` alone — see KNOWN LIMITATIONS above.
    getChanges: (): FolioAgentChange[] => [],
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
};
