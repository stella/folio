import { Fragment, type Node as PMNode } from "prosemirror-model";

type InlineBookmarksForParagraphJoinOptions = {
  first: PMNode;
  boundaries: readonly PMNode[];
  second: PMNode;
};

/** The inline form of boundaries whose enclosing paragraph break was removed. */
export const inlineBookmarksOf = (boundaries: readonly PMNode[]): PMNode[] | null => {
  const inlineType = boundaries.at(0)?.type.schema.nodes["bookmarkBoundary"];
  if (!inlineType || boundaries.some((node) => node.type.name !== "blockBookmarkBoundary")) {
    return null;
  }
  return boundaries.map((node) => inlineType.create(node.attrs));
};

/** Inline boundaries that can replace the paragraph break without replacing either paragraph's content. */
export const inlineBookmarksForParagraphJoin = ({
  first,
  boundaries,
  second,
}: InlineBookmarksForParagraphJoinOptions): PMNode[] | null => {
  if (first.type.name !== "paragraph" || second.type !== first.type) return null;
  const inline = inlineBookmarksOf(boundaries);
  if (!inline) return null;
  const content = Fragment.fromArray([
    ...first.content.content,
    ...inline,
    ...second.content.content,
  ]);
  return first.type.validContent(content) ? inline : null;
};
