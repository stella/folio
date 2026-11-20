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

import { setDocumentBodyComments } from "@stll/docx-core/model";

import {
  type CommentIdAllocator,
  createCommentIdAllocator,
} from "../prosemirror/commentIdAllocator";
import type { Comment, Document } from "../types/document";
import { deterministicHexId } from "../utils/hexId";
import { commentTextParagraph } from "./commentTextParagraph";

export type CreateCommentReplyInput = {
  author: string;
  text: string;
  initials?: string;
  /** ISO date; defaults to now. */
  date?: string;
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

/**
 * A paragraph id no comment paragraph holds, derived from `seed`, so the same
 * reply to the same comments gets the same id.
 */
const freshParaId = (used: Set<string>, seed: string): string => {
  let attempt = 0;
  let id = deterministicHexId(seed);
  while (used.has(id.toUpperCase())) {
    attempt += 1;
    id = deterministicHexId(`${seed}:${String(attempt)}`);
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
export type CreateReplyOptions = {
  comments: readonly Comment[];
  parentCommentId: number;
  input: CreateCommentReplyInput;
  /**
   * The document's comment ids. The reply's id is the next of them, above
   * every comment in `comments`.
   */
  commentIds: CommentIdAllocator;
};

export const createReply = ({
  comments,
  parentCommentId,
  input,
  commentIds,
}: CreateReplyOptions): Comment | null => {
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
  const threadRoot =
    comments.find((comment) => comment.id === threadRootId) ?? parent;
  const rootLastParagraph = (threadRoot.content ?? []).at(-1);
  if (!rootLastParagraph) {
    return null;
  }

  for (const { id } of comments) {
    commentIds.seedAbove(id);
  }
  const used = usedParaIds(comments);
  // The root's last paragraph must carry a paraId so the reply can reference it
  // via `w15:paraIdParent`; a doc authored without paraIds gets one now.
  if (!rootLastParagraph.paraId) {
    rootLastParagraph.paraId = freshParaId(
      used,
      `comment:${String(threadRootId)}`,
    );
  }

  const id = commentIds.next();
  const replyParagraph = commentTextParagraph(input.text);
  replyParagraph.paraId = freshParaId(used, `comment:${String(id)}`);

  return {
    id,
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
  // The document's comments are its comment ids: the reply takes the next.
  const reply = createReply({
    comments,
    parentCommentId,
    input,
    commentIds: createCommentIdAllocator(),
  });
  if (!reply) {
    return null;
  }
  setDocumentBodyComments(doc.package.document, [...comments, reply]);
  return reply;
};
