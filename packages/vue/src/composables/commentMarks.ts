/**
 * ProseMirror comment-mark helpers for the Vue adapter's add-comment flow.
 *
 * Core owns the paraId/search-based `addCommentToRange` (commentOps.ts) and the
 * id allocator, but the interactive "select text → pending highlight → type →
 * submit" flow needs range-based mark ops that core does not expose. These
 * mirror the React adapter's `commentsHelpers.ts` / `commentAnchors.ts` so both
 * adapters apply the same pending/real comment marks and resolve the same
 * creation range.
 */

import { Selection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { PENDING_COMMENT_ID } from "@stll/folio-core/prosemirror/commentIdAllocator";
import { findBodyPmAnchors } from "@stll/folio-core/layout-bridge/dom/findBodyPmSpans";

export type CommentMarkRange = {
  from: number;
  to: number;
};

/** Clamp a range into the document and reject an empty (collapsed) span. */
export function clampCommentMarkRange(
  docSize: number,
  range: CommentMarkRange,
): CommentMarkRange | null {
  const from = Math.max(0, Math.min(range.from, docSize));
  const to = Math.max(from, Math.min(range.to, docSize));
  if (from === to) return null;
  return { from, to };
}

/**
 * Pick the range to comment on, preferring the range captured when the FAB was
 * shown, then the live selection, then the last non-empty selection. Mirrors
 * React's `resolveCommentCreationRange`.
 */
export type ResolveCommentCreationRangeOptions = {
  docSize: number;
  capturedRange: CommentMarkRange;
  currentRange: CommentMarkRange;
  savedRange: CommentMarkRange | null;
};

export function resolveCommentCreationRange({
  docSize,
  capturedRange,
  currentRange,
  savedRange,
}: ResolveCommentCreationRangeOptions): CommentMarkRange | null {
  const range = (() => {
    if (capturedRange.from !== capturedRange.to) return capturedRange;
    if (currentRange.from !== currentRange.to) return currentRange;
    if (savedRange && savedRange.from !== savedRange.to) return savedRange;
    return null;
  })();
  return range ? clampCommentMarkRange(docSize, range) : null;
}

export type ApplyCommentMarkOptions = {
  replacePending?: boolean;
  selectEnd?: boolean;
};

/**
 * Add a `comment` mark for `commentId` over `range`. With `replacePending` the
 * placeholder mark (PENDING_COMMENT_ID) is removed from the range first so a
 * pre-existing real comment the selection overlaps is left intact. Returns
 * false when the schema lacks the mark or the range is empty.
 */
export function applyCommentMarkRange(
  view: EditorView,
  range: CommentMarkRange,
  commentId: number,
  options?: ApplyCommentMarkOptions,
): boolean {
  const commentMark = view.state.schema.marks["comment"];
  const safeRange = clampCommentMarkRange(view.state.doc.content.size, range);
  if (!commentMark || !safeRange) return false;

  let tr = view.state.tr;
  if (options?.replacePending) {
    tr = tr.removeMark(
      safeRange.from,
      safeRange.to,
      commentMark.create({ commentId: PENDING_COMMENT_ID }),
    );
  }
  tr = tr.addMark(safeRange.from, safeRange.to, commentMark.create({ commentId }));

  if (options?.selectEnd) {
    tr = tr.setSelection(Selection.near(tr.doc.resolve(safeRange.to), -1));
  }

  view.dispatch(tr);
  return true;
}

/** Strip the pending placeholder comment mark from `range` (cancel add-comment). */
export function removePendingCommentMarkRange(view: EditorView, range: CommentMarkRange): void {
  const commentMark = view.state.schema.marks["comment"];
  const safeRange = clampCommentMarkRange(view.state.doc.content.size, range);
  if (!commentMark || !safeRange) return;
  view.dispatch(
    view.state.tr.removeMark(
      safeRange.from,
      safeRange.to,
      commentMark.create({ commentId: PENDING_COMMENT_ID }),
    ),
  );
}

const FALLBACK_TOP_OFFSET = 80;

/**
 * Y position (relative to `scrollContainer`) of the painted element covering
 * PM position `pmPos`, using the same `data-pm-start`/`data-pm-end` anchors the
 * React adapter reads. Returns null when no anchor covers the position yet.
 */
export function findSelectionYPosition(
  scrollContainer: HTMLElement | null,
  pagesContainer: HTMLElement | null,
  pmPos: number,
): number | null {
  if (!scrollContainer || !pagesContainer) return null;
  for (const el of findBodyPmAnchors(pagesContainer)) {
    const pmStart = Number(el.dataset["pmStart"]);
    const pmEnd = Number(el.dataset["pmEnd"]);
    if (pmPos >= pmStart && pmPos <= pmEnd) {
      return (
        el.getBoundingClientRect().top -
        scrollContainer.getBoundingClientRect().top +
        scrollContainer.scrollTop
      );
    }
  }
  return null;
}

export function getFallbackCommentYPosition(scrollContainer: HTMLElement | null): number {
  if (!scrollContainer) return FALLBACK_TOP_OFFSET;
  return scrollContainer.scrollTop + Math.max(FALLBACK_TOP_OFFSET, scrollContainer.clientHeight / 3);
}
