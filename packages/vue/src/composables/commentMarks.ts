/**
 * The range the Vue adapter's add-comment flow comments on. The highlight and
 * the thread's marks come from core's comment records; this picks the range,
 * as React's `commentAnchors.ts` does.
 */

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
