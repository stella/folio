/**
 * Create a threaded reply to an existing comment.
 *
 * Produces a new `Comment` whose `parentId` points at the target comment and
 * whose last paragraph carries a fresh `w14:paraId`, and guarantees the parent
 * has a paraId too — the two ids Word uses to link a reply to its thread in
 * `commentsExtended.xml` (`w15:paraId` / `w15:paraIdParent`). The reply's
 * `document.xml` range markers are synthesized at save time by
 * {@link applyReplyThreadMarkers}, so no anchor work is needed here.
 */

import type { Comment, Document, Paragraph } from "../types/document";
import { generateHexId } from "../utils/hexId";

export type CreateCommentReplyInput = {
  author: string;
  text: string;
  initials?: string;
  /** ISO date; defaults to now. */
  date?: string;
};

const nextCommentId = (comments: readonly Comment[]): number => {
  let max = 0;
  for (const comment of comments) {
    if (comment.id > max) {
      max = comment.id;
    }
  }
  return max + 1;
};

const usedParaIds = (comments: readonly Comment[]): Set<string> => {
  const ids = new Set<string>();
  for (const comment of comments) {
    for (const paragraph of comment.content) {
      if (paragraph.paraId) {
        ids.add(paragraph.paraId.toUpperCase());
      }
    }
  }
  return ids;
};

const freshParaId = (used: Set<string>): string => {
  let id = generateHexId();
  while (used.has(id.toUpperCase())) {
    id = generateHexId();
  }
  used.add(id.toUpperCase());
  return id;
};

/**
 * Build a reply to `parentCommentId` against an existing comment list and
 * return it, mutating the thread root's last paragraph to guarantee it has a
 * paraId. Does NOT append the reply to `comments` — the caller owns storage
 * (the standalone {@link replyToComment} writes it back to the document model;
 * the headless reviewer keeps its own created-comment list). Returns `null` when
 * the parent comment does not exist, or is malformed (no paragraphs) so it
 * cannot anchor a thread. A reply to a reply re-parents onto the same thread as
 * the target (Word threads are flat: every reply links to the root).
 */
export const createReply = (
  comments: readonly Comment[],
  parentCommentId: number,
  input: CreateCommentReplyInput,
): Comment | null => {
  const parent = comments.find((comment) => comment.id === parentCommentId);
  if (!parent) {
    return null;
  }
  // Word models replies as a flat list under the thread root, so a reply to a
  // reply attaches to that reply's parent rather than nesting.
  const threadRootId = parent.parentId ?? parent.id;

  // The root's last paragraph is the thread key; a comment with no paragraphs
  // (malformed / empty) cannot anchor a reply, so refuse rather than index into
  // an empty array.
  const threadRoot = comments.find((comment) => comment.id === threadRootId) ?? parent;
  const rootLastParagraph = (threadRoot.content ?? []).at(-1);
  if (!rootLastParagraph) {
    return null;
  }

  const used = usedParaIds(comments);
  // The root's last paragraph must carry a paraId so the reply can reference it
  // via `w15:paraIdParent`; a doc authored without paraIds gets one now.
  if (!rootLastParagraph.paraId) {
    rootLastParagraph.paraId = freshParaId(used);
  }

  const replyParagraph: Paragraph = {
    type: "paragraph",
    formatting: {},
    paraId: freshParaId(used),
    content: [{ type: "run", formatting: {}, content: [{ type: "text", text: input.text }] }],
  };

  return {
    id: nextCommentId(comments),
    author: input.author,
    date: input.date ?? new Date().toISOString(),
    parentId: threadRootId,
    content: [replyParagraph],
    ...(input.initials !== undefined ? { initials: input.initials } : {}),
  };
};

/**
 * Append a reply to `parentCommentId` in `doc`'s comment list, mutating the
 * document model in place, and return the created reply. Returns `null` when the
 * parent comment does not exist or cannot anchor a thread (see
 * {@link createReply}).
 */
export const replyToComment = (
  doc: Document,
  parentCommentId: number,
  input: CreateCommentReplyInput,
): Comment | null => {
  const comments = doc.package.document.comments ?? [];
  const reply = createReply(comments, parentCommentId, input);
  if (!reply) {
    return null;
  }
  doc.package.document.comments = [...comments, reply];
  return reply;
};
