import type {
  FolioAIEditApplyMode,
  FolioDocxReviewer,
  FolioReviewChange,
  FolioReviewComment,
  FolioReviewCommentReply,
} from "@stll/folio-core/server";

import type { FolioAgentBridge } from "../bridge";
import type { FolioAgentChange, FolioAgentComment, FolioAgentCommentReply } from "../types";

/** Options for {@link createReviewerBridge}. */
export type CreateReviewerBridgeOptions = {
  /** `"tracked-changes"` (default) produces ins/del redlines; `"direct"` edits in place. */
  mode?: FolioAIEditApplyMode;
};

const toAgentCommentReply = (reply: FolioReviewCommentReply): FolioAgentCommentReply => ({
  id: String(reply.id),
  author: reply.author,
  text: reply.text,
});

const toAgentComment = (comment: FolioReviewComment): FolioAgentComment => ({
  id: String(comment.id),
  author: comment.author,
  text: comment.text,
  resolved: comment.done,
  blockId: comment.blockId,
  quote: comment.anchoredText,
  replies: comment.replies.map(toAgentCommentReply),
});

const toAgentChange = (change: FolioReviewChange): FolioAgentChange => ({
  id: String(change.id),
  type: change.type,
  author: change.author,
  text: change.text,
  blockId: change.blockId,
});

/**
 * Build a {@link FolioAgentBridge} over a headless {@link FolioDocxReviewer}
 * (`@stll/folio-core/server`). No optional capability members are
 * implemented: a headless document has no live page/selection/scroll
 * surface, so `read_page`, `read_selection`, and `scroll_to_block` report an
 * unsupported-capability error via {@link executeFolioToolCall} on this
 * bridge.
 */
export const createReviewerBridge = (
  reviewer: FolioDocxReviewer,
  options: CreateReviewerBridgeOptions = {},
): FolioAgentBridge => {
  const mode = options.mode ?? "tracked-changes";

  return {
    snapshot: () => reviewer.snapshot(),
    applyOperations: (operations) => reviewer.applyOperations(operations, { mode }),
    getComments: () => reviewer.getComments().map(toAgentComment),
    getChanges: () => reviewer.getChanges().map(toAgentChange),
    replyToComment: (commentId, text) => {
      const parentId = Number.parseInt(commentId, 10);
      if (Number.isNaN(parentId)) {
        return false;
      }
      return reviewer.replyTo(parentId, { text }) !== null;
    },
    resolveComment: (commentId, resolved) => reviewer.resolveComment(commentId, { resolved }),
  };
};
