import { Mark, type Node as PMNode } from "prosemirror-model";

import { INLINE_WRAPPER_MARK_NAME } from "./extensions/marks/InlineWrapperExtension";
import { getInlineDisplayText } from "./inlineDisplayText";
import { RUN_IDENTITY_MARK_NAME } from "./runIdentity";

/** Current children alone decide whether a field needs editable structural content. */
export const fieldRequiresStructuredContent = (
  children: readonly PMNode[],
): boolean => {
  const first = children.at(0);
  return children.some(
    (child) =>
      !child.isText ||
      (first !== undefined && !Mark.sameSet(first.marks, child.marks)) ||
      child.marks.some(
        (mark) =>
          mark.type.name === RUN_IDENTITY_MARK_NAME ||
          mark.type.name === "hyperlink" ||
          mark.type.name === INLINE_WRAPPER_MARK_NAME ||
          mark.type.name === "insertion" ||
          mark.type.name === "deletion",
      ),
  );
};

/** Resolution collapses a field only when the atom can retain its result. */
export const canonicalFieldNode = (node: PMNode): PMNode => {
  if (node.type.name !== "structuredField") return node;
  const displayText = getInlineDisplayText(node);
  const attrs = { ...node.attrs, displayText };
  if (fieldRequiresStructuredContent(node.content.content)) {
    return node.attrs["displayText"] === displayText
      ? node
      : node.type.create(attrs, node.content, node.marks);
  }
  let marks = node.marks;
  for (const mark of node.firstChild?.marks ?? []) marks = mark.addToSet(marks);
  return node.type.schema.node("field", attrs, undefined, marks);
};
