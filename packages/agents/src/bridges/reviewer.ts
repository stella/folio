import type {
  FolioAIEditApplyMode,
  FolioDocxReviewer,
  FolioReviewComment,
  FolioReviewCommentReply,
} from "@stll/folio-core/server";

import type { FolioAgentBridge } from "../bridge";
import { registerDecodedCommentHandlers } from "../bridge";
import { decodeCommentId } from "../codecs";
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

  const bridge: FolioAgentBridge = {
    snapshot: () => reviewer.snapshot(),
    documentOperationMode: mode,
    applyDocumentOperations: (batch) =>
      reviewer.applyDocumentOperations(batch.mode === undefined ? { ...batch, mode } : batch),
    undoDocumentOperations: (undoHandle) => reviewer.undoDocumentOperations(undoHandle),
    getComments: () => reviewer.getComments().map(toAgentComment),
    getChanges: () => reviewer.getChanges().map(toAgentChange),
    listStories: () => reviewer.listStories(),
    readStory: (handle) => reviewer.readStory(handle),
    replyToComment: (commentId, text) => {
      const decoded = decodeCommentId(commentId);
      return decoded !== null && reviewer.replyTo(decoded.numeric, { text }) !== null;
    },
    resolveComment: (commentId, resolved) => {
      const decoded = decodeCommentId(commentId);
      return decoded !== null && reviewer.resolveComment(decoded.raw, { resolved });
    },
  };
  return registerDecodedCommentHandlers(bridge, {
    replyToComment: ({ numeric }, text) => reviewer.replyTo(numeric, { text }) !== null,
    resolveComment: ({ raw }, resolved) => reviewer.resolveComment(raw, { resolved }),
  });
};
