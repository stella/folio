/**
 * Bring a painted page's content up to a fresh rendering of it while keeping
 * the painted nodes that did not change. Replacing the whole content area for
 * one edited paragraph makes the browser restyle, lay out, paint and raster
 * every line on the page; keeping the unchanged nodes limits that to the
 * paragraphs the edit touched.
 */

import { releaseDisplayPageResources } from "../display-list/pageResourceOwnership";

/**
 * The attributes that say where a painted node stands in the editor document.
 * An edit before a paragraph shifts them without changing what it shows, so a
 * node that differs from its fresh rendering only in these is kept and takes
 * the fresh values.
 */
const POSITION_ATTRIBUTES: ReadonlySet<string> = new Set([
  "data-pm-start",
  "data-pm-end",
  "data-table-pm-start",
  "data-sdt-pm-pos",
]);

const sameAttributes = (
  painted: Element,
  fresh: Element,
  ignoring: ReadonlySet<string>,
): boolean => {
  if (painted.attributes.length !== fresh.attributes.length) return false;
  for (const attribute of Array.from(fresh.attributes)) {
    if (ignoring.has(attribute.name)) {
      if (!painted.hasAttribute(attribute.name)) return false;
      continue;
    }
    if (painted.getAttribute(attribute.name) !== attribute.value) return false;
  }
  return true;
};

const NO_ATTRIBUTES: ReadonlySet<string> = new Set();

/** Whether `painted` shows what `fresh` shows, positions aside. */
const showsTheSame = (painted: Node, fresh: Node): boolean => {
  if (painted.nodeType !== fresh.nodeType) return false;
  if (painted.nodeType !== painted.ELEMENT_NODE)
    return painted.isEqualNode(fresh);
  const paintedElement = painted as Element;
  const freshElement = fresh as Element;
  if (
    paintedElement.namespaceURI !== freshElement.namespaceURI ||
    paintedElement.localName !== freshElement.localName ||
    !sameAttributes(paintedElement, freshElement, POSITION_ATTRIBUTES)
  )
    return false;
  const paintedChildren = painted.childNodes;
  const freshChildren = fresh.childNodes;
  if (paintedChildren.length !== freshChildren.length) return false;
  for (let index = 0; index < freshChildren.length; index += 1) {
    const paintedChild = paintedChildren[index];
    const freshChild = freshChildren[index];
    if (!paintedChild || !freshChild || !showsTheSame(paintedChild, freshChild))
      return false;
  }
  return true;
};

/** Give `painted` the position attributes of `fresh`, which shows the same. */
const takePositions = (painted: Node, fresh: Node): void => {
  if (painted.nodeType !== painted.ELEMENT_NODE) return;
  const paintedElement = painted as Element;
  const freshElement = fresh as Element;
  for (const name of POSITION_ATTRIBUTES) {
    const value = freshElement.getAttribute(name);
    if (value !== null && paintedElement.getAttribute(name) !== value)
      paintedElement.setAttribute(name, value);
  }
  const paintedChildren = painted.childNodes;
  const freshChildren = fresh.childNodes;
  for (let index = 0; index < freshChildren.length; index += 1) {
    const paintedChild = paintedChildren[index];
    const freshChild = freshChildren[index];
    if (paintedChild && freshChild) takePositions(paintedChild, freshChild);
  }
};

/**
 * Make `painted` equal to `fresh` (as `isEqualNode` compares them), keeping
 * every child of `painted` that shows what its counterpart in `fresh` shows
 * and moving the others in from `fresh`. Returns false, changing nothing, when
 * the two differ in their own attributes or in their number of children: the
 * caller then replaces `painted` whole.
 *
 * The fresh children it moves leave `fresh`; what remains there is the
 * caller's to release.
 */
export const reconcileRenderedContent = (
  painted: Element,
  fresh: Element,
): boolean => {
  if (
    painted.localName !== fresh.localName ||
    !sameAttributes(painted, fresh, NO_ATTRIBUTES) ||
    painted.childNodes.length !== fresh.childNodes.length
  )
    return false;
  const paintedChildren = Array.from(painted.childNodes);
  const freshChildren = Array.from(fresh.childNodes);
  for (const [index, freshChild] of freshChildren.entries()) {
    const paintedChild = paintedChildren[index];
    if (!paintedChild) return false;
    if (paintedChild.isEqualNode(freshChild)) continue;
    if (showsTheSame(paintedChild, freshChild)) {
      takePositions(paintedChild, freshChild);
      continue;
    }
    if (paintedChild.nodeType === paintedChild.ELEMENT_NODE)
      releaseDisplayPageResources(paintedChild as Element);
    paintedChild.replaceWith(freshChild);
  }
  return true;
};
