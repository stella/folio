import type {
  FolioAIEditApplyMode,
  FolioDocxReviewer,
  FolioRevisionStamp,
  FolioReviewComment,
  FolioReviewCommentReply,
} from "@stll/folio-core/server";
import type { FolioHeadlessPageSource } from "@stll/folio-core/reviewer-page-source";

import type { FolioAgentBridge } from "../bridge";
import { registerDecodedCommentHandlers, registerOperationSession } from "../bridge";
import { decodeCommentId } from "../codecs";
import type { FolioAgentComment, FolioAgentCommentReply } from "../types";
import { toAgentChange } from "./shared";

/** Options for {@link createReviewerBridge}. */
export type CreateReviewerBridgeOptions = {
  /** `"tracked-changes"` (default) produces ins/del redlines; `"direct"` edits in place. */
  mode?: FolioAIEditApplyMode;
  /**
   * Fixed provenance for what this bridge authors: revisions and comments
   * take its date and allocate ids from its seed, and replies take its date.
   * For a bridge that applies one batch (a file transaction); a second batch
   * would allocate from the same seed. Omit to use the wall clock.
   */
  revisionStamp?: FolioRevisionStamp;
  /** Optional pagination over the reviewer's current headless document state. */
  pageSource?: FolioHeadlessPageSource;
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
 * (`@stll/folio-core/server`). Page capabilities are attached only when a
 * caller supplies a headless page source; selection and scrolling remain
 * live-editor-only capabilities.
 */
export const createReviewerBridge = (
  reviewer: FolioDocxReviewer,
  options: CreateReviewerBridgeOptions = {},
): FolioAgentBridge => {
  const mode = options.mode ?? "tracked-changes";
  const { revisionStamp } = options;
  const applyOptions = revisionStamp === undefined ? {} : { revisionStamp };
  const replyDate = revisionStamp === undefined ? {} : { date: revisionStamp.date };

  const bridge: FolioAgentBridge = {
    snapshot: () => reviewer.snapshot(),
    documentOperationMode: mode,
    applyDocumentOperations: (batch) =>
      reviewer.applyDocumentOperations(batch.mode === undefined ? { ...batch, mode } : batch, {
        ...applyOptions,
        undefinedReferences: "refuse",
      }),
    undoDocumentOperations: (undoHandle) => reviewer.undoDocumentOperations(undoHandle),
    getComments: () => reviewer.getComments().map(toAgentComment),
    getChanges: () => reviewer.getChanges().map(toAgentChange),
    listStories: () => reviewer.listStories(),
    readStory: (handle) => reviewer.readStory(handle),
    replyToComment: (commentId, text) => {
      const decoded = decodeCommentId(commentId);
      return decoded !== null && reviewer.replyTo(decoded.numeric, { text, ...replyDate }) !== null;
    },
    resolveComment: (commentId, resolved) => {
      const decoded = decodeCommentId(commentId);
      return decoded !== null && reviewer.resolveComment(decoded.raw, { resolved });
    },
  };
  const { pageSource } = options;
  if (pageSource) {
    bridge.getPageCount = () => pageSource.pageCount();
    bridge.getPageText = (page) => pageSource.pageText(page);
    bridge.getTargetPage = (target) => pageSource.targetPage(target);
  }
  registerOperationSession(bridge, reviewer);
  return registerDecodedCommentHandlers(bridge, {
    replyToComment: ({ numeric }, text) =>
      reviewer.replyTo(numeric, { text, ...replyDate }) !== null,
    resolveComment: ({ raw }, resolved) => reviewer.resolveComment(raw, { resolved }),
  });
};
