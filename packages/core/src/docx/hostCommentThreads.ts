/**
 * The comment threads an editor host keeps beside the editor, written into the
 * document it saves or hands out.
 *
 * Hosts keep the threads (text, author, replies, resolution) in their own
 * list, which comment edits change without touching the ProseMirror
 * document. A document built from the editor carries the threads it was
 * loaded with, so every host writes its list in through this one rule.
 */

import type { Comment } from "../types/content";
import type { Document } from "../types/document";
import { setDocumentBodyComments } from "./canonicalModel";

export function getCommentParentId(
  comment: Comment,
): number | null | undefined {
  const runtimeComment: { parentId?: number | null } = comment;
  return runtimeComment.parentId;
}

/**
 * The comments whose reply chain reaches a top-level comment.
 *
 * Every top-level comment is kept, whether or not its anchor survives:
 * documents without reply-thread metadata deliver replies as top-level
 * comments, and dropping unanchored ones would drop those replies. Replies
 * whose parent chain does not reach a kept top-level comment are dropped.
 */
export function pruneOrphanedComments(
  comments: readonly Comment[],
  _referencedIds?: ReadonlySet<number>,
): Comment[] {
  const keptIds = new Set<number>();
  for (const comment of comments) {
    const parentId = getCommentParentId(comment);
    if (parentId === null || parentId === undefined) {
      keptIds.add(comment.id);
    }
  }
  // Pull in replies whose parent is kept until no more join, so a thread of
  // depth N is complete after N - 1 passes.
  let addedThisPass = true;
  while (addedThisPass) {
    addedThisPass = false;
    for (const comment of comments) {
      if (keptIds.has(comment.id)) {
        continue;
      }
      const parentId = getCommentParentId(comment);
      if (
        parentId !== null &&
        parentId !== undefined &&
        keptIds.has(parentId)
      ) {
        keptIds.add(comment.id);
        addedThisPass = true;
      }
    }
  }
  return comments.filter((comment) => keptIds.has(comment.id));
}

/**
 * `document` with its body's comment threads replaced by the host's. The body
 * and package are copied, so the document passed in is left as it was.
 */
export function withHostCommentThreads(
  document: Document,
  comments: readonly Comment[],
): Document {
  const body = { ...document.package.document };
  setDocumentBodyComments(body, pruneOrphanedComments(comments));
  return { ...document, package: { ...document.package, document: body } };
}
