import { Fragment, type Node as PMNode } from "prosemirror-model";

import { recreateProseNodeWithParagraphPropertySource } from "../../docx/paragraphPropertySource";

type JoinParagraphsAcrossBookmarksOptions = {
  first: PMNode;
  boundaries: readonly PMNode[];
  second: PMNode;
  owner: PMNode;
  attrs: PMNode["attrs"];
};

/** The inline form of boundaries whose enclosing paragraph break was removed. */
export const inlineBookmarksOf = (boundaries: readonly PMNode[]): PMNode[] | null => {
  const inlineType = boundaries.at(0)?.type.schema.nodes["bookmarkBoundary"];
  if (!inlineType || boundaries.some((node) => node.type.name !== "blockBookmarkBoundary")) {
    return null;
  }
  return boundaries.map((node) => inlineType.create(node.attrs));
};

/** Place block-level bookmark boundaries at the text junction when a break goes away. */
export const joinParagraphsAcrossBookmarks = ({
  first,
  boundaries,
  second,
  owner,
  attrs,
}: JoinParagraphsAcrossBookmarksOptions): PMNode | null => {
  if (first.type.name !== "paragraph" || second.type !== first.type || owner.type !== first.type)
    return null;
  const inline = inlineBookmarksOf(boundaries);
  if (!inline) return null;
  const content = Fragment.fromArray([
    ...first.content.content,
    ...inline,
    ...second.content.content,
  ]);
  return first.type.validContent(content)
    ? recreateProseNodeWithParagraphPropertySource(owner, { attrs, content })
    : null;
};
