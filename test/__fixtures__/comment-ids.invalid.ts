/** Deliberate violations: comment ids tested for truth, which skips comment 0. */

type CommentMark = { attrs: { commentId: number } };

export const paintedIds = (marks: readonly CommentMark[]): number[] => {
  const ids: number[] = [];
  for (const mark of marks) {
    const { commentId } = mark.attrs;
    if (commentId) ids.push(commentId);
  }
  return ids;
};

export const activeThread = (activeCommentId: number | null): string => {
  if (!activeCommentId) return "none";
  return String(activeCommentId);
};

export const firstMarkId = (marks: readonly CommentMark[]): number | null =>
  marks[0]?.attrs.commentId ? marks[0].attrs.commentId : null;

export const hasAnchor = (mark: CommentMark, painted: boolean): boolean =>
  Boolean(mark.attrs.commentId && painted);
