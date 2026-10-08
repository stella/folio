import type { Node as PMNode } from "prosemirror-model";
import { INLINE_WRAPPER_MARK_NAME } from "./extensions/marks/InlineWrapperExtension";

/** Current children alone decide whether a field needs editable structural content. */
export const fieldRequiresStructuredContent = (children: readonly PMNode[]): boolean =>
  children.some(
    (child) =>
      child.type.name === "pageBreakRun" ||
      child.type.name === "preservedXml" ||
      child.marks.some(
        (mark) =>
          mark.type.name === "hyperlink" ||
          mark.type.name === INLINE_WRAPPER_MARK_NAME ||
          mark.type.name === "insertion" ||
          mark.type.name === "deletion",
      ),
  );

/** Resolution returns run-only fields to the same atom representation as initial conversion. */
export const canonicalFieldNode = (node: PMNode): PMNode => {
  if (node.type.name !== "structuredField") return node;
  const attrs = { ...node.attrs, displayText: node.textContent };
  if (fieldRequiresStructuredContent(node.content.content)) {
    return node.attrs["displayText"] === node.textContent
      ? node
      : node.type.create(attrs, node.content, node.marks);
  }
  let marks = node.marks;
  for (const mark of node.firstChild?.marks ?? []) marks = mark.addToSet(marks);
  return node.type.schema.node("field", attrs, undefined, marks);
};
