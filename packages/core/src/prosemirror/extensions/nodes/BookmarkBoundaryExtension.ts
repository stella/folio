/** Zero-width bookmark boundary that preserves its position through ProseMirror edits. */

import { expectBookmarkBoundaryAttrs } from "../../bookmarkBoundaryAttrs";
import { createNodeExtension } from "../create";

type BookmarkBoundaryOptions = {
  getInternalClipboardToken?: () => string;
};

function readNonnegativeInteger(value: string | null): number | false {
  if (value === null || !/^(?:0|[1-9]\d*)$/.test(value)) {
    return false;
  }
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : false;
}

function readOptionalColumn(dom: HTMLElement, attribute: string): number | undefined | false {
  const value = dom.getAttribute(attribute);
  return value === null ? undefined : readNonnegativeInteger(value);
}

export const BookmarkBoundaryExtension = createNodeExtension<BookmarkBoundaryOptions>({
  name: "bookmarkBoundary",
  schemaNodeName: "bookmarkBoundary",
  nodeSpec: (options) => ({
    inline: true,
    group: "inline",
    marks: "_",
    atom: true,
    selectable: false,
    attrs: {
      type: {},
      id: {},
      name: { default: null },
      colFirst: { default: null },
      colLast: { default: null },
    },
    parseDOM: [
      {
        tag: "span[data-docx-bookmark-boundary]",
        getAttrs(dom) {
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
          return {
            type,
            id,
            ...(name ? { name } : {}),
            ...(colFirst !== undefined ? { colFirst } : {}),
            ...(colLast !== undefined ? { colLast } : {}),
          };
        },
      },
    ],
    toDOM(node) {
      const attrs = expectBookmarkBoundaryAttrs(node);
      return [
        "span",
        {
          "data-docx-bookmark-boundary": attrs.type,
          "data-docx-bookmark-id": String(attrs.id),
          ...(attrs.type === "start" ? { "data-docx-bookmark-name": attrs.name } : {}),
          ...(attrs.type === "start" && attrs.colFirst !== undefined
            ? { "data-docx-bookmark-col-first": String(attrs.colFirst) }
            : {}),
          ...(attrs.type === "start" && attrs.colLast !== undefined
            ? { "data-docx-bookmark-col-last": String(attrs.colLast) }
            : {}),
          "aria-hidden": "true",
          contenteditable: "false",
          style: "display: none;",
          ...(options.getInternalClipboardToken
            ? { "data-docx-internal-clipboard": options.getInternalClipboardToken() }
            : {}),
        },
      ];
    },
  }),
});
