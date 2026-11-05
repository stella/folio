/**
 * How an inline node with no visible content sits in the editor's DOM.
 *
 * It takes no space, but it keeps a box. A browser's editing commands skip what
 * has no box, so replacing the composition text beside a `display: none` element
 * also removed the element, and the next DOM read lost the node it stood for.
 */
export const INLINE_ANCHOR_DOM_STYLE =
  "display: inline-block; width: 0; height: 0; overflow: hidden;";
