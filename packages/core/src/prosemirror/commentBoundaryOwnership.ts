export type TrackedCommentBoundary = {
  readonly type: "start" | "end";
  readonly commentId: number;
  readonly owner: "revision" | "parent";
};

type TrackedCommentBoundaryOwnerOptions = {
  boundaries: readonly TrackedCommentBoundary[] | null | undefined;
  type: TrackedCommentBoundary["type"];
  commentId: number;
  fallback: TrackedCommentBoundary["owner"];
};

export const trackedCommentBoundaryOwner = ({
  boundaries,
  type,
  commentId,
  fallback,
}: TrackedCommentBoundaryOwnerOptions): TrackedCommentBoundary["owner"] =>
  boundaries?.find((boundary) => boundary.type === type && boundary.commentId === commentId)
    ?.owner ?? fallback;
