/**
 * useCommentManagement — comment thread state + mutation handlers for the Vue
 * adapter, wired to core's comment ops. Mirrors the React adapter's
 * useFolioComments + the sidebar mutation callbacks in DocxEditor.tsx, adapted
 * to our core primitives:
 *   - id allocation:      commentIdAllocator (createCommentIdAllocator, seed)
 *   - comment factory:    commentOps.createComment(allocator, ...)
 *   - accept / reject:    revision-scoped acceptAIEditRevision / rejectAIEditRevision
 *                         (by-id path) or range-based acceptChange / rejectChange
 *                         (raw selection path)
 *
 * Controlled (`comments` prop set) vs uncontrolled: when controlled, the thread
 * list is read from the prop and every mutation only routes through
 * `onCommentsChange`; when uncontrolled it also updates internal state and seeds
 * from the loaded document.
 */

import { computed, ref, type ComputedRef, type Ref } from "vue";
import type { EditorView } from "prosemirror-view";
import type { Comment } from "@stll/folio-core/types/content";
import type { Document } from "@stll/folio-core/types/document";
import {
  createCommentIdAllocator,
  seedCommentAllocator,
} from "@stll/folio-core/prosemirror/commentIdAllocator";
import { createComment as coreCreateComment } from "@stll/folio-core/prosemirror/commentOps";
import {
  acceptChange,
  rejectChange,
  acceptAIEditRevision,
  rejectAIEditRevision,
} from "@stll/folio-core/prosemirror/commands/comments";

export type UseCommentManagementOptions = {
  editorView: Ref<EditorView | null>;
  getDocument: () => Document | null;
  author: () => string;
  /** Reader for the controlled `comments` prop (undefined => uncontrolled). */
  commentsProp: () => Comment[] | undefined;
  /** Fires on every comment-list mutation (controlled and uncontrolled). */
  onCommentsChange: (comments: Comment[]) => void;
  /** Re-run layout after a mutation that changed marks in the document. */
  reLayout: () => void;
};

export type UseCommentManagementReturn = {
  comments: ComputedRef<Comment[]>;
  /**
   * True when the comment thread list has user mutations not yet serialized by
   * `save()`. Mirrors React's `commentsDirtyRef`: comment edits (add / reply /
   * resolve / unresolve / delete) do not touch the PM document, so they are
   * invisible to the doc-dirty flag; this surfaces them. Set on every user
   * comment mutation, cleared on document load (`seedFromDocument`) and on a
   * successful save (`clearCommentsDirty`).
   */
  commentsDirty: ComputedRef<boolean>;
  /** Clear {@link commentsDirty} after a successful save. */
  clearCommentsDirty: () => void;
  /**
   * Build a comment with a freshly-allocated, collision-seeded id.
   * `authorOverride` lets a caller (AI-edit operations) mint a comment
   * attributed to the resolved operation author instead of the editor's
   * default `options.author()` — mirrors React's `applyAIEditOperations`
   * closure, which creates each comment with the per-call `operationAuthor`.
   */
  createComment: (text: string, parentId?: number, authorOverride?: string) => Comment;
  /** Append a comment and emit the change. */
  pushComment: (comment: Comment) => void;
  /** Replace the whole list and emit the change (uncontrolled also stores it). */
  setComments: (next: Comment[]) => void;
  /** Seed internal comments + the id allocator from the loaded document. */
  seedFromDocument: () => void;
  handleReply: (parentId: number, text: string) => void;
  handleResolve: (commentId: number) => void;
  handleUnresolve: (commentId: number) => void;
  handleDelete: (commentId: number) => void;
  handleTrackedChangeReply: (revisionId: number, text: string) => void;
  handleAcceptChangeById: (revisionId: number) => void;
  handleRejectChangeById: (revisionId: number) => void;
  handleAcceptChange: (from: number, to: number) => void;
  handleRejectChange: (from: number, to: number) => void;
};

export function useCommentManagement(
  options: UseCommentManagementOptions,
): UseCommentManagementReturn {
  const allocator = createCommentIdAllocator();
  const internalComments = ref<Comment[]>([]);
  // Dirty flag for comment-only edits (see the return type doc). Not set inside
  // `setComments`, which is also the load-seed path — only the user-driven
  // mutation handlers below mark it, matching React (where the low-level comment
  // setter is clean and only `replaceComments` sets `commentsDirtyRef`).
  const commentsDirty = ref(false);

  function markCommentsDirty(): void {
    commentsDirty.value = true;
  }

  function clearCommentsDirty(): void {
    commentsDirty.value = false;
  }

  const isControlled = computed(() => options.commentsProp() !== undefined);
  const comments = computed<Comment[]>(() =>
    isControlled.value ? (options.commentsProp() ?? []) : internalComments.value,
  );

  function setComments(next: Comment[]): void {
    if (!isControlled.value) internalComments.value = next;
    options.onCommentsChange(next);
  }

  function pushComment(comment: Comment): void {
    markCommentsDirty();
    setComments([...comments.value, comment]);
  }

  function createComment(text: string, parentId?: number, authorOverride?: string): Comment {
    // Raise the allocator above every id currently live in the model + doc so a
    // reply/new comment can never re-mint an existing comment or revision id.
    seedCommentAllocator(allocator, comments.value, options.editorView.value);
    return coreCreateComment(allocator, text, authorOverride ?? options.author(), parentId);
  }

  function seedFromDocument(): void {
    // A document load/swap is a clean baseline: the seeded comments came from the
    // just-loaded file, so nothing is pending (mirrors React clearing
    // `commentsDirtyRef` on document reset). Cleared unconditionally so a
    // controlled host's swap resets the flag too.
    clearCommentsDirty();
    if (isControlled.value) return;
    const bodyComments = options.getDocument()?.package.document.comments;
    seedCommentAllocator(allocator, bodyComments, options.editorView.value);
    // Reset to the loaded document's comments (empty when it has none) so a
    // document swap does not leak the previous document's comments into the
    // sidebar. Matches React, which clears comments on document reset.
    setComments(bodyComments && bodyComments.length > 0 ? bodyComments : []);
  }

  function handleReply(parentId: number, text: string): void {
    pushComment(createComment(text, parentId));
  }

  function handleResolve(commentId: number): void {
    markCommentsDirty();
    setComments(comments.value.map((c) => (c.id === commentId ? { ...c, done: true } : c)));
  }

  function handleUnresolve(commentId: number): void {
    markCommentsDirty();
    setComments(comments.value.map((c) => (c.id === commentId ? { ...c, done: false } : c)));
  }

  function handleDelete(commentId: number): void {
    // Drop the comment and any direct replies threaded under it (matches React's
    // onCommentDelete). The orphaned comment mark left in the doc is pruned at
    // save time.
    markCommentsDirty();
    setComments(comments.value.filter((c) => c.id !== commentId && c.parentId !== commentId));
  }

  function handleTrackedChangeReply(revisionId: number, text: string): void {
    pushComment(createComment(text, revisionId));
  }

  function resolveChangeById(revisionId: number, accept: boolean): void {
    const view = options.editorView.value;
    if (!view) return;
    // Revision-scoped: acceptAIEditRevision/rejectAIEditRevision both locate
    // the range for this specific revisionId AND thread it through as the
    // match set, so only marks/structural changes carrying this id resolve.
    // A plain range-based acceptChange(from, to)/rejectChange(from, to) here
    // would treat the missing revisionId as "match everything in range,"
    // which can also resolve an unrelated collaborator's tracked change
    // sharing the same paragraph/row span.
    const command = accept ? acceptAIEditRevision : rejectAIEditRevision;
    command(revisionId)(view.state, view.dispatch);
  }

  function handleAcceptChangeById(revisionId: number): void {
    resolveChangeById(revisionId, true);
  }

  function handleRejectChangeById(revisionId: number): void {
    resolveChangeById(revisionId, false);
  }

  function handleAcceptChange(from: number, to: number): void {
    const view = options.editorView.value;
    if (!view) return;
    acceptChange(from, to)(view.state, view.dispatch);
  }

  function handleRejectChange(from: number, to: number): void {
    const view = options.editorView.value;
    if (!view) return;
    rejectChange(from, to)(view.state, view.dispatch);
  }

  return {
    comments,
    commentsDirty: computed(() => commentsDirty.value),
    clearCommentsDirty,
    createComment,
    pushComment,
    setComments,
    seedFromDocument,
    handleReply,
    handleResolve,
    handleUnresolve,
    handleDelete,
    handleTrackedChangeReply,
    handleAcceptChangeById,
    handleRejectChangeById,
    handleAcceptChange,
    handleRejectChange,
  };
}
