import { Fragment, type Node as PMNode } from "prosemirror-model";
import { canJoin } from "prosemirror-transform";

import {
  inlineBoundariesOf,
  isBlockBoundaryNode,
} from "../keptRangeBoundaries";

type InlineBookmarksForParagraphJoinOptions = {
  first: PMNode;
  boundaries: readonly PMNode[];
  second: PMNode;
};

/** The inline form of boundaries whose enclosing paragraph break was removed. */
export const inlineBookmarksOf = (
  boundaries: readonly PMNode[],
): PMNode[] | null => inlineBoundariesOf(boundaries);

/** Inline boundaries that can replace the paragraph break without replacing either paragraph's content. */
export const inlineBookmarksForParagraphJoin = ({
  first,
  boundaries,
  second,
}: InlineBookmarksForParagraphJoinOptions): PMNode[] | null => {
  if (first.type.name !== "paragraph" || second.type !== first.type)
    return null;
  const inline = inlineBookmarksOf(boundaries);
  if (!inline) return null;
  const content = Fragment.fromArray([
    ...first.content.content,
    ...inline,
    ...second.content.content,
  ]);
  return first.type.validContent(content) ? inline : null;
};

/** Paragraph siblings separated only by typed range boundaries. */
export const paragraphJoinAfter = (
  doc: PMNode,
  position: number,
): {
  next: PMNode;
  nextPosition: number;
  inlineBookmarks: PMNode[];
} | null => {
  const at = doc.resolve(position);
  const first = at.nodeBefore;
  if (first?.type.name !== "paragraph") return null;
  const boundaries: PMNode[] = [];
  let nextPosition = position;
  let index = at.index();
  while (index < at.parent.childCount) {
    const boundary = at.parent.child(index);
    if (!isBlockBoundaryNode(boundary)) break;
    boundaries.push(boundary);
    nextPosition += boundary.nodeSize;
    index++;
  }
  const next = index < at.parent.childCount ? at.parent.child(index) : null;
  if (!next || next.type !== first.type) return null;
  if (boundaries.length === 0) {
    return canJoin(doc, position)
      ? { next, nextPosition, inlineBookmarks: [] }
      : null;
  }
  const inlineBookmarks = inlineBookmarksForParagraphJoin({
    first,
    boundaries,
    second: next,
  });
  return inlineBookmarks ? { next, nextPosition, inlineBookmarks } : null;
};
