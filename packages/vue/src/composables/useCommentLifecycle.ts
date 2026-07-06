/**
 * useCommentLifecycle — the interactive "add comment" flow for the Vue adapter:
 * the floating add-comment button that tracks a non-empty selection, the pending
 * comment highlight while the input is open, and submit / cancel. Mirrors the
 * React adapter's floatingCommentBtn + onAddComment / onCancelAddComment, built
 * on the shared comment-mark helpers (commentMarks.ts) and core's id allocator
 * (via the injected `createComment`).
 */

import { ref, type Ref } from "vue";
import type { EditorView } from "prosemirror-view";
import type { Comment } from "@stll/folio-core/types/content";
import { PENDING_COMMENT_ID } from "@stll/folio-core/prosemirror/commentIdAllocator";
import {
  applyCommentMarkRange,
  removePendingCommentMarkRange,
  resolveCommentCreationRange,
  type CommentMarkRange,
} from "./commentMarks";

/** Right edge fallback offset when no page element is measurable. */
const PAGE_RIGHT_FALLBACK_OFFSET = 408;
const FAB_GAP = 12;

export type FloatingCommentButton = {
  top: number;
  left: number;
  from: number;
  to: number;
};

export type UseCommentLifecycleOptions = {
  editorView: Ref<EditorView | null>;
  /** The absolute-positioned viewport the FAB is rendered into. */
  pagesViewport: Ref<HTMLElement | null>;
  /** The painted `.paged-editor__pages` container (anchor coordinate space). */
  pagesContainer: Ref<HTMLElement | null>;
  readOnly: Ref<boolean>;
  /** Build a comment with a freshly-allocated id (from useCommentManagement). */
  createComment: (text: string, parentId?: number) => Comment;
  /** Append a comment + emit the change (from useCommentManagement). */
  pushComment: (comment: Comment) => void;
  reLayout: () => void;
  showSidebar: Ref<boolean>;
  /** Highlight the freshly-added comment card in the sidebar. */
  setActiveSidebarItem: (id: string | null) => void;
};

export type UseCommentLifecycleReturn = {
  isAddingComment: Ref<boolean>;
  addCommentYPosition: Ref<number | null>;
  floatingCommentButton: Ref<FloatingCommentButton | null>;
  /** Recompute the floating button from the current selection (call on select). */
  updateFloatingButton: () => void;
  /** Enter the add-comment flow for the captured / current selection. */
  startAddComment: () => void;
  /** Submit the typed comment text; returns false when it could not be applied. */
  handleAddComment: (text: string) => boolean;
  /** Abandon the in-progress comment and clear the pending highlight. */
  handleCancelAddComment: () => void;
};

export function useCommentLifecycle(
  options: UseCommentLifecycleOptions,
): UseCommentLifecycleReturn {
  const isAddingComment = ref(false);
  const addCommentYPosition = ref<number | null>(null);
  const floatingCommentButton = ref<FloatingCommentButton | null>(null);
  const commentSelectionRange = ref<CommentMarkRange | null>(null);
  // Last non-empty selection, so a submit still resolves a range after the
  // selection collapsed (clicking the input, re-layout).
  const lastSelectionRange = ref<CommentMarkRange | null>(null);

  /** Y of a PM position relative to the painted pages container (card space). */
  function anchorYInPages(pos: number): number | null {
    const view = options.editorView.value;
    const pages = options.pagesContainer.value;
    if (!view || !pages) return null;
    try {
      return view.coordsAtPos(pos).top - pages.getBoundingClientRect().top;
    } catch {
      return null;
    }
  }

  function updateFloatingButton(): void {
    const view = options.editorView.value;
    const viewport = options.pagesViewport.value;
    if (!view || !viewport || isAddingComment.value || options.readOnly.value) {
      floatingCommentButton.value = null;
      return;
    }
    const { from, to } = view.state.selection;
    if (from === to) {
      floatingCommentButton.value = null;
      return;
    }
    lastSelectionRange.value = { from, to };

    let coords: { top: number };
    try {
      coords = view.coordsAtPos(from);
    } catch {
      floatingCommentButton.value = null;
      return;
    }
    const viewportRect = viewport.getBoundingClientRect();
    const pageRect = options.pagesContainer.value?.getBoundingClientRect();
    const rawLeft = pageRect
      ? pageRect.right - viewportRect.left + FAB_GAP
      : viewportRect.width / 2 + PAGE_RIGHT_FALLBACK_OFFSET;
    floatingCommentButton.value = {
      top: coords.top - viewportRect.top,
      left: Math.max(16, Math.min(rawLeft, viewportRect.width - 16)),
      from,
      to,
    };
  }

  function startAddComment(): void {
    const view = options.editorView.value;
    if (!view) return;
    const button = floatingCommentButton.value;
    const selection = view.state.selection;
    const safeRange = resolveCommentCreationRange({
      docSize: view.state.doc.content.size,
      capturedRange: button ? { from: button.from, to: button.to } : { from: 0, to: 0 },
      currentRange: { from: selection.from, to: selection.to },
      savedRange: lastSelectionRange.value,
    });
    if (!safeRange) {
      floatingCommentButton.value = null;
      return;
    }
    if (!applyCommentMarkRange(view, safeRange, PENDING_COMMENT_ID, { selectEnd: true })) {
      floatingCommentButton.value = null;
      commentSelectionRange.value = null;
      return;
    }
    commentSelectionRange.value = safeRange;
    addCommentYPosition.value = anchorYInPages(safeRange.from) ?? button?.top ?? null;
    options.showSidebar.value = true;
    isAddingComment.value = true;
    floatingCommentButton.value = null;
  }

  function resetAddCommentState(): void {
    isAddingComment.value = false;
    commentSelectionRange.value = null;
    addCommentYPosition.value = null;
  }

  function handleAddComment(text: string): boolean {
    const view = options.editorView.value;
    const range = commentSelectionRange.value;
    if (!view || !range) return false;

    const comment = options.createComment(text);
    if (!applyCommentMarkRange(view, range, comment.id, { replacePending: true })) return false;

    options.pushComment(comment);
    options.reLayout();
    options.setActiveSidebarItem(`comment-${comment.id}`);
    resetAddCommentState();
    return true;
  }

  function handleCancelAddComment(): void {
    const view = options.editorView.value;
    if (view && commentSelectionRange.value) {
      removePendingCommentMarkRange(view, commentSelectionRange.value);
    }
    resetAddCommentState();
  }

  return {
    isAddingComment,
    addCommentYPosition,
    floatingCommentButton,
    updateFloatingButton,
    startAddComment,
    handleAddComment,
    handleCancelAddComment,
  };
}
