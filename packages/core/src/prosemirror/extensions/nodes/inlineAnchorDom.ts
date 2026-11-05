import type { Node as PMNode } from "prosemirror-model";

/**
 * How an inline node with no visible content sits in the editor's DOM.
 *
 * It takes no space, but it keeps a box. A browser's editing commands skip what
 * has no box, so replacing the composition text beside a `display: none` element
 * also removed the element, and the next DOM read lost the node it stood for.
 */
export const INLINE_ANCHOR_DOM_STYLE =
  "display: inline-block; width: 0; height: 0; overflow: hidden;";

const INLINE_ANCHOR_NODE_NAMES: ReadonlySet<string> = new Set([
  "bookmarkBoundary",
  "commentRangeBoundary",
  "commentReference",
  "moveRangeBoundary",
  "noteMarker",
  "rangeAnchor",
  "renderedPageBreak",
]);

/** Whether a node is one of those, which the caret passes without stopping. */
export const isInlineAnchor = (node: PMNode): boolean =>
  INLINE_ANCHOR_NODE_NAMES.has(node.type.name);
