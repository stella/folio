/**
 * What a bookmark boundary looks like in the DOM, at either level.
 *
 * The inline atom and the block one are two ProseMirror node types because
 * ProseMirror decides inline or block per type, not per node. Everything else
 * about them is the same marker, so the attribute spec, the DOM it writes and
 * the DOM it reads back live here once: a bookmark cannot come back describing
 * itself differently depending on which side of a paragraph it was written on.
 */

import type { Node as PMNode } from "prosemirror-model";

import { expectBookmarkBoundaryAttrs, isDisplacedByCustomXml } from "../../bookmarkBoundaryAttrs";

const readNonnegativeInteger = (value: string | null): number | false => {
  if (value === null || !/^(?:0|[1-9]\d*)$/.test(value)) {
    return false;
  }
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : false;
};

const readOptionalColumn = (dom: HTMLElement, attribute: string): number | undefined | false => {
  const value = dom.getAttribute(attribute);
  return value === null ? undefined : readNonnegativeInteger(value);
};

export const bookmarkBoundaryAttrSpec = {
  type: {},
  id: {},
  name: { default: null },
  colFirst: { default: null },
  colLast: { default: null },
  displacedByCustomXml: { default: null },
};

/** `false` when the element does not describe a boundary ProseMirror can hold. */
export const parseBookmarkBoundaryDom = (
  dom: HTMLElement | string,
): Record<string, unknown> | false => {
  if (typeof dom === "string") {
    return false;
  }
  const type = dom.getAttribute("data-docx-bookmark-boundary");
  const id = readNonnegativeInteger(dom.getAttribute("data-docx-bookmark-id"));
  if ((type !== "start" && type !== "end") || id === false) {
    return false;
  }
  const name = dom.getAttribute("data-docx-bookmark-name");
  if (type === "start" && !name) {
    return false;
  }
  const colFirst = readOptionalColumn(dom, "data-docx-bookmark-col-first");
  const colLast = readOptionalColumn(dom, "data-docx-bookmark-col-last");
  if (colFirst === false || colLast === false) {
    return false;
  }
  const displaced = dom.getAttribute("data-docx-bookmark-displaced");
  return {
    type,
    id,
    ...(name ? { name } : {}),
    ...(colFirst !== undefined ? { colFirst } : {}),
    ...(colLast !== undefined ? { colLast } : {}),
    ...(isDisplacedByCustomXml(displaced) ? { displacedByCustomXml: displaced } : {}),
  };
};

export const bookmarkBoundaryDomAttributes = (
  node: PMNode,
  internalClipboardToken?: string,
): Record<string, string> => {
  const attrs = expectBookmarkBoundaryAttrs(node);
  return {
    "data-docx-bookmark-boundary": attrs.type,
    "data-docx-bookmark-id": String(attrs.id),
    ...(attrs.type === "start" ? { "data-docx-bookmark-name": attrs.name } : {}),
    ...(attrs.type === "start" && attrs.colFirst !== undefined
      ? { "data-docx-bookmark-col-first": String(attrs.colFirst) }
      : {}),
    ...(attrs.type === "start" && attrs.colLast !== undefined
      ? { "data-docx-bookmark-col-last": String(attrs.colLast) }
      : {}),
    ...(attrs.displacedByCustomXml !== undefined
      ? { "data-docx-bookmark-displaced": attrs.displacedByCustomXml }
      : {}),
    "aria-hidden": "true",
    contenteditable: "false",
    style: "display: none;",
    ...(internalClipboardToken === undefined
      ? {}
      : { "data-docx-internal-clipboard": internalClipboardToken }),
  };
};
