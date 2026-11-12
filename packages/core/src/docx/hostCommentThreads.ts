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
import type { BlockContent, Document } from "../types/document";
import { setDocumentBodyComments } from "./canonicalModel";
import { visitCommentMarkers } from "./commentAnchorIndex";

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

/** The comment ids the saved stories of `document` anchor. */
const commentIdsAnchoredIn = (document: Document): Set<number> => {
  const ids = new Set<number>();
  const visit = (blocks: readonly BlockContent[]) =>
    visitCommentMarkers(blocks, ({ item }) => {
      ids.add(item.id);
    });
  const { package: pkg } = document;
  visit(pkg.document.content);
  for (const story of [
    ...(pkg.headers?.values() ?? []),
    ...(pkg.footers?.values() ?? []),
    ...(pkg.footnotes ?? []),
    ...(pkg.endnotes ?? []),
  ])
    visit(story.content);
  return ids;
};

export type HostCommentThreadsOptions = {
  /**
   * The comment ids the editor's own document anchors, or `null` when the
   * host has no editor document. A thread anchored there that the document
   * being saved leaves out (a comment a pending suggestion proposes, which
   * the save holds back with the suggestion) is left out with its replies:
   * a saved thread without its anchor is a stray card on reopening.
   */
  anchoredInEditor: ReadonlySet<number> | null;
};

/**
 * `document` with its body's comment threads replaced by the host's. The body
 * and package are copied, so the document passed in is left as it was.
 */
export function withHostCommentThreads(
  document: Document,
  comments: readonly Comment[],
  { anchoredInEditor }: HostCommentThreadsOptions,
): Document {
  const body = { ...document.package.document };
  const kept = (() => {
    if (anchoredInEditor === null || anchoredInEditor.size === 0)
      return comments;
    const saved = commentIdsAnchoredIn(document);
    const heldBack = new Set(
      comments
        .filter(({ id }) => anchoredInEditor.has(id) && !saved.has(id))
        .map(({ id }) => id),
    );
    return heldBack.size === 0
      ? comments
      : comments.filter(({ id }) => !heldBack.has(id));
  })();
  // Replies under a thread left out lose their parent, and pruning drops them.
  setDocumentBodyComments(body, pruneOrphanedComments(kept));
  return { ...document, package: { ...document.package, document: body } };
}
