// A field-totality gate in a test file is read by no compiler, so the review
// comment field census lives here, where the package typecheck compiles it.
import type { FolioReviewComment, FolioReviewCommentReply } from "../headless";

export const REVIEW_COMMENT_FIELD_CENSUS: Record<
  keyof FolioReviewComment,
  keyof FolioReviewComment
> = {
  id: "id",
  author: "author",
  date: "date",
  text: "text",
  anchoredText: "anchoredText",
  blockId: "blockId",
  story: "story",
  replies: "replies",
  done: "done",
};

export const REVIEW_COMMENT_REPLY_FIELD_CENSUS: Record<
  keyof FolioReviewCommentReply,
  keyof FolioReviewCommentReply
> = {
  id: "id",
  author: "author",
  date: "date",
  text: "text",
};
