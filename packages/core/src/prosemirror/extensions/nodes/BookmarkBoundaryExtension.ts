/** Zero-width bookmark boundary that preserves its position through ProseMirror edits. */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";

import { readParagraphAttrs } from "../../attrs";
import {
  expectBookmarkBoundaryAttrs,
  readBookmarkBoundaryAttrs,
} from "../../bookmarkBoundaryAttrs";
import {
  findInvalidBookmarkBoundaryIds,
  type BookmarkBoundaryOccurrence,
} from "../../bookmarkBoundaryIntegrity";
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

/** The node name, for callers asking whether a paragraph holds any content. */
export const BOOKMARK_BOUNDARY_NODE_NAME = "bookmarkBoundary";

type PositionedNode = {
  position: number;
  node: PMNode;
};

type PositionedBoundary = BookmarkBoundaryOccurrence & PositionedNode;

const collectInvalidBoundaries = (doc: PMNode): PositionedNode[] => {
  const boundaries: PositionedBoundary[] = [];
  const malformedBoundaries: PositionedNode[] = [];
  const paragraphBookmarkIds = new Set<number>();

  doc.descendants((node, position) => {
    if (node.type.name === "paragraph") {
      const attrs = readParagraphAttrs(node);
      if (attrs.ok) {
        for (const bookmark of attrs.value.bookmarks ?? []) {
          paragraphBookmarkIds.add(bookmark.id);
        }
      }
    }
    if (node.type.name !== BOOKMARK_BOUNDARY_NODE_NAME) {
      return true;
    }

    const result = readBookmarkBoundaryAttrs(node);
    if (!result.ok) {
      malformedBoundaries.push({ position, node });
      return false;
    }
    const attrs = result.value;
    const boundary = { id: attrs.id, type: attrs.type, position, node };
    boundaries.push(boundary);
    return false;
  });

  const invalidIds = findInvalidBookmarkBoundaryIds(boundaries, paragraphBookmarkIds);
  return [...boundaries.filter(({ id }) => invalidIds.has(id)), ...malformedBoundaries];
};

export const BookmarkBoundaryExtension = createNodeExtension<BookmarkBoundaryOptions>({
  name: BOOKMARK_BOUNDARY_NODE_NAME,
  schemaNodeName: BOOKMARK_BOUNDARY_NODE_NAME,
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
  onSchemaReady: () => ({
    plugins: [
      new Plugin({
        appendTransaction(transactions, _oldState, newState) {
          if (!transactions.some(({ docChanged }) => docChanged)) {
            return null;
          }

          const invalidBoundaries = collectInvalidBoundaries(newState.doc);
          if (invalidBoundaries.length === 0) {
            return null;
          }

          const transaction = newState.tr;
          for (const { node, position } of invalidBoundaries.toSorted(
            (first, second) => second.position - first.position,
          )) {
            transaction.delete(position, position + node.nodeSize);
          }
          return transaction;
        },
      }),
    ],
  }),
});
