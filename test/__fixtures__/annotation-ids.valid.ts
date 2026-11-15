/** Annotation ids tested for absence, lists tested for length, named ids. */

type CommentMark = { attrs: { commentId: number } };

export const paintedIds = (marks: readonly CommentMark[]): number[] => {
  const ids: number[] = [];
  for (const mark of marks) ids.push(mark.attrs.commentId);
  return ids;
};

export const activeThread = (activeCommentId: number | null): string =>
  activeCommentId === null ? "none" : String(activeCommentId);

export const hasComments = (commentIds: readonly number[]): boolean =>
  commentIds.length > 0;

export const fallbackId = (commentId: number | undefined): number =>
  commentId ?? -1;

export const styleName = (styleId: string | undefined): string =>
  styleId ? styleId : "Normal";
