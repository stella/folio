import type {
  FolioAIEditApplyMode,
  FolioDocxReviewer,
  FolioReviewComment,
  FolioReviewCommentReply,
} from "@stll/folio-core/server";

import type { FolioAgentBridge } from "../bridge";
import type { FolioAgentComment, FolioAgentCommentReply } from "../types";
import { toAgentChange } from "./shared";

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

/**
 * Parse a tool-supplied comment id into the numeric id `FolioDocxReviewer`
 * expects, rejecting anything but a bare non-negative integer. `Number.
 * parseInt` alone would silently accept trailing junk (`"12abc"` -> `12`),
 * which could reply to the wrong comment thread on malformed tool input.
 */
const parseCommentId = (commentId: string): number | null =>
  /^\d+$/.test(commentId) ? Number.parseInt(commentId, 10) : null;

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
    applyDocumentOperations: (batch) => reviewer.applyDocumentOperations({ ...batch, mode }),
    getComments: () => reviewer.getComments().map(toAgentComment),
    getChanges: () => reviewer.getChanges().map(toAgentChange),
    replyToComment: (commentId, text) => {
      const parentId = parseCommentId(commentId);
      if (parentId === null) {
        return false;
      }
      return reviewer.replyTo(parentId, { text }) !== null;
    },
    resolveComment: (commentId, resolved) => reviewer.resolveComment(commentId, { resolved }),
  };
};
