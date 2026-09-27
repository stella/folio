/**
 * Which comment threads an edit has left with nothing to anchor to.
 *
 * Word removes a comment together with the content it is anchored to. The
 * `w:commentReference` run is part of that content (ECMA-376 Part 1
 * §17.13.4.5), so rejecting the insertion that holds it, accepting the
 * deletion of it or deleting it outright takes the reference, and a
 * `comments.xml` definition (§17.13.4.2) that nothing references anymore has
 * no place in the document to be shown. The editor already drops the orphan
 * reference and the marks go with their text; the definition has to follow,
 * or the reviewer lists a thread about nothing and the save writes it back.
 *
 * Only a comment whose anchor an edit removed is lost. A definition the
 * source package already left unanchored is kept as it was: a reply in a
 * file without `commentsExtended.xml` looks exactly like one, and writing
 * back what the edit never touched is the conservative choice. Measuring
 * "had an anchor" and "has one now" with the same reading of each story
 * also keeps anchors folio does not project (captured markup) out of it.
 */

import type { Node as PMNode } from "prosemirror-model";

import { visitCommentMarkers } from "../docx/commentAnchorIndex";
import { readCommentReferenceAttrs } from "../prosemirror/commentReferenceAttrs";
import { COMMENT_REFERENCE_NODE_NAME } from "../prosemirror/extensions/nodes/CommentReferenceExtension";
import { RANGE_ANCHOR_NODE_NAME } from "../prosemirror/extensions/nodes/RangeAnchorExtension";
import { readRangeAnchorAttrs } from "../prosemirror/rangeAnchorAttrs";
import type { BlockContent, Comment } from "../types/document";

/** Ids of the comments an editor document anchors: by mark, reference or point anchor. */
export const anchoredCommentIdsInProseDoc = (doc: PMNode): Set<number> => {
  const ids = new Set<number>();
  doc.descendants((node) => {
    if (node.type.name === COMMENT_REFERENCE_NODE_NAME) {
      const attrs = readCommentReferenceAttrs(node);
      if (attrs.ok) {
        ids.add(attrs.value.commentId);
      }
    } else if (node.type.name === RANGE_ANCHOR_NODE_NAME) {
      const attrs = readRangeAnchorAttrs(node);
      if (attrs.ok && attrs.value.start.type === "commentRangeStart") {
        ids.add(attrs.value.start.id);
      }
    }
    for (const mark of node.marks) {
      const commentId = mark.attrs["commentId"];
      if (mark.type.name === "comment" && typeof commentId === "number") {
        ids.add(commentId);
      }
    }
    return true;
  });
  return ids;
};

/** Ids of the comments a model story's blocks hold any marker of. */
export const anchoredCommentIdsInBlocks = (blocks: readonly BlockContent[]): Set<number> => {
  const ids = new Set<number>();
  visitCommentMarkers(blocks, ({ item }) => ids.add(item.id));
  return ids;
};

/**
 * `comments` without the threads whose root lost its anchor: the root, and
 * every reply that threads under it. A reply's own markers repeat its root's
 * range, so a reply is judged by its root, never on its own.
 */
export const withoutLostCommentThreads = (
  comments: readonly Comment[],
  {
    anchoredBefore,
    anchoredNow,
  }: { anchoredBefore: ReadonlySet<number>; anchoredNow: ReadonlySet<number> },
): Comment[] => {
  const dropped = new Set<number>();
  for (const comment of comments) {
    if (
      comment.parentId === undefined &&
      anchoredBefore.has(comment.id) &&
      !anchoredNow.has(comment.id)
    ) {
      dropped.add(comment.id);
    }
  }
  if (dropped.size === 0) {
    return [...comments];
  }
  let grew = true;
  while (grew) {
    grew = false;
    for (const comment of comments) {
      if (
        !dropped.has(comment.id) &&
        comment.parentId !== undefined &&
        dropped.has(comment.parentId)
      ) {
        dropped.add(comment.id);
        grew = true;
      }
    }
  }
  return comments.filter((comment) => !dropped.has(comment.id));
};
