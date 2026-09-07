import type { Node as PMNode, ResolvedPos } from "prosemirror-model";

/**
 * A node the FILE holds as a child of the container it appears in.
 *
 * The editor keeps a text box as a block-level sibling of the paragraph that
 * anchors it; the file hangs it off that paragraph's own run, so it is not a
 * child of the container at all. A paragraph and a table are children in both.
 */
const isAContainerChild = (node: PMNode, paragraphTypeName: string): boolean =>
  node.type.name === paragraphTypeName || node.type.spec["tableRole"] === "table";

/**
 * Whether the paragraph starting at `at` is the one its container ENDS with,
 * which is the one place a deleted paragraph mark cannot say what it means.
 *
 * A deleted mark asks a reader to join the paragraph it ends with the one
 * after it. A body, a table cell, a header or footer, a note and a text box
 * each end with a paragraph that nothing follows, so the ask cannot be carried
 * out and a consumer refuses the whole package rather than opening it.
 *
 * A paragraph before a TABLE is not one of those: the table is a following
 * sibling, so the paragraph does not end anything and carries its own mark
 * like any other.
 */
export const paragraphEndsItsContainer = (at: ResolvedPos, paragraphTypeName: string): boolean => {
  for (let index = at.index() + 1; index < at.parent.childCount; index++) {
    if (isAContainerChild(at.parent.child(index), paragraphTypeName)) {
      return false;
    }
  }
  return true;
};
