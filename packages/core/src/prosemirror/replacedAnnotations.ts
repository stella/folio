/**
 * What survives when an edit replaces a span of inline content.
 *
 * ProseMirror gives a replacement the marks it inherits at the boundary, and
 * every annotation mark in this schema is `inclusive: false`, so a span that
 * reaches the end of what a `comment` covers loses that comment. For
 * formatting that is right; for an annotation it is not. The comment's range
 * started on the replaced text, so the save emits a `commentRangeEnd` with no
 * `commentRangeStart`: invalid OOXML, and a comment anchored to nothing.
 * Zero-width anchors fare worse: a plain replace deletes them outright.
 *
 * Word moves a range start that sits inside a replaced span to the beginning
 * of the replacement, and lets the range keep covering the new text. A range
 * wholly inside the span collapses onto the replacement, and disappears only
 * when every character it covered is gone. One rule per marker kind, and that
 * is the whole contract:
 *
 * - `comment` mark: the replacement carries every comment the span covered, so
 *   a range that started inside it starts at the replacement instead.
 * - `commentReference` node: kept after the replacement, where the comment's
 *   last covered position now is.
 * - `bookmarkBoundary` node: a start kept before the replacement, an end after
 *   it, so the bookmark still covers the new text.
 * - `textBoxAnchor` node: kept after the replacement; the shape it anchors
 *   outlives the sentence it was typed into.
 * - `renderedPageBreak` node: dropped with the text. It is Word's cached
 *   pagination, re-derived on layout, and nothing else names it.
 * - tracked-change marks (`insertion`, `deletion`, `moveFrom`, `moveTo`): not
 *   carried. The applier decides them per edit mode and writes them itself,
 *   and a zero-width anchor may carry no revision of its own at all.
 * - formatting marks: not touched. Boundary inheritance already gives the
 *   replacement the formatting of the text it replaces.
 *
 * Deleting the whole of a comment's span leaves it no mark anywhere, so
 * `CommentReferenceExtension` drops the orphan reference and the save drops
 * the comment: the one case where a range is meant to disappear.
 */

import { Fragment, Mark, type Node as PMNode } from "prosemirror-model";

import { BOOKMARK_BOUNDARY_NODE_NAME } from "./extensions/nodes/BookmarkBoundaryExtension";
import { COMMENT_REFERENCE_NODE_NAME } from "./extensions/nodes/CommentReferenceExtension";
import { TEXT_BOX_ANCHOR_NODE_NAME } from "./extensions/nodes/TextBoxAnchorExtension";

const COMMENT_MARK_NAME = "comment";

/** Anchors an edit must put back, and where relative to the replacement. */
const ANCHOR_PLACEMENT = {
  [COMMENT_REFERENCE_NODE_NAME]: "after",
  [TEXT_BOX_ANCHOR_NODE_NAME]: "after",
  [BOOKMARK_BOUNDARY_NODE_NAME]: "byBoundaryType",
} as const;

type AnchorPlacement = (typeof ANCHOR_PLACEMENT)[keyof typeof ANCHOR_PLACEMENT];

const anchorPlacement = (node: PMNode): AnchorPlacement | undefined =>
  ANCHOR_PLACEMENT[node.type.name as keyof typeof ANCHOR_PLACEMENT];

export type ReplacedAnnotations = {
  /** Comment marks the replacement must carry, in document order. */
  comments: readonly Mark[];
  /** Anchors restored before the replacement, in their original order. */
  leading: readonly PMNode[];
  /** Anchors restored after the replacement, in their original order. */
  trailing: readonly PMNode[];
};

const NOTHING_REPLACED: ReplacedAnnotations = { comments: [], leading: [], trailing: [] };

export const hasReplacedAnnotations = ({
  comments,
  leading,
  trailing,
}: ReplacedAnnotations): boolean =>
  comments.length > 0 || leading.length > 0 || trailing.length > 0;

/**
 * The annotations carried by `[from, to)`. An anchor counts only when the
 * range covers it whole: a partially covered atom is not deleted by the
 * replacement either.
 */
export const surveyReplacedAnnotations = (
  doc: PMNode,
  from: number,
  to: number,
): ReplacedAnnotations => {
  if (to <= from) {
    return NOTHING_REPLACED;
  }

  const comments: Mark[] = [];
  const seenCommentIds = new Set<unknown>();
  const leading: PMNode[] = [];
  const trailing: PMNode[] = [];

  doc.nodesBetween(from, to, (node, position) => {
    if (!node.isInline) {
      return true;
    }
    for (const mark of node.marks) {
      if (mark.type.name !== COMMENT_MARK_NAME || seenCommentIds.has(mark.attrs["commentId"])) {
        continue;
      }
      seenCommentIds.add(mark.attrs["commentId"]);
      comments.push(mark);
    }
    const placement = anchorPlacement(node);
    if (placement === undefined || position < from || position + node.nodeSize > to) {
      return false;
    }
    const opensARange = placement === "byBoundaryType" && node.attrs["type"] === "start";
    (opensARange ? leading : trailing).push(node);
    return false;
  });

  return { comments, leading, trailing };
};

type AnnotatedReplacementOptions = {
  annotations: ReplacedAnnotations;
  /** The replacement's own inline content, already carrying its formatting. */
  content: Fragment;
};

type AnnotatedReplacement = {
  fragment: Fragment;
  /** Size of the anchors restored before the replacement's own content. */
  leadingSize: number;
  /** Size of the replacement's own content, once the comments are on it. */
  contentSize: number;
};

const anchorsSize = (anchors: readonly PMNode[]): number => {
  let size = 0;
  for (const node of anchors) {
    size += node.nodeSize;
  }
  return size;
};

const withComments = (node: PMNode, comments: readonly Mark[]): PMNode => {
  let marks = node.marks;
  for (const mark of comments) {
    marks = mark.addToSet(marks);
  }
  return node.mark(marks);
};

/** `content` with the surveyed comments on it and the surveyed anchors around it. */
export const annotatedReplacement = ({
  annotations: { comments, leading, trailing },
  content,
}: AnnotatedReplacementOptions): AnnotatedReplacement => {
  const inline: PMNode[] = [];
  content.forEach((child) => {
    inline.push(comments.length === 0 ? child : withComments(child, comments));
  });

  const fragment = Fragment.fromArray([...leading, ...inline, ...trailing]);
  const leadingSize = anchorsSize(leading);
  return {
    fragment,
    leadingSize,
    contentSize: fragment.size - leadingSize - anchorsSize(trailing),
  };
};

/**
 * The marks a replacement inherits at `[from, to)`, exactly as
 * `Transaction.insertText` would compute them.
 */
export const inheritedReplacementMarks = (
  doc: PMNode,
  from: number,
  to: number,
): readonly Mark[] => {
  const $from = doc.resolve(from);
  return (to === from ? $from.marks() : $from.marksAcross(doc.resolve(to))) ?? Mark.none;
};
