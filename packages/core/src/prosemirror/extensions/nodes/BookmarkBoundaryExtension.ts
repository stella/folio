/** Zero-width bookmark boundary that preserves its position through ProseMirror edits. */

import type { Node as PMNode } from "prosemirror-model";

import type { PositionedBookmarkMarker } from "../../../types/document";
import { Plugin } from "prosemirror-state";

import { readParagraphAttrs, readTableAttrs, readTableRowAttrs } from "../../attrs";
import { readBookmarkBoundaryAttrs } from "../../bookmarkBoundaryAttrs";
import {
  findInvalidBookmarkBoundaryIds,
  type BookmarkBoundaryOccurrence,
} from "../../bookmarkBoundaryIntegrity";
import { createNodeExtension } from "../create";

import { BLOCK_BOOKMARK_BOUNDARY_NODE_NAME } from "./BlockBookmarkBoundaryExtension";
import {
  bookmarkBoundaryAttrSpec,
  bookmarkBoundaryDomAttributes,
  parseBookmarkBoundaryDom,
} from "./bookmarkBoundaryDom";

type BookmarkBoundaryOptions = {
  getInternalClipboardToken?: () => string;
};

/** The node name, for callers asking whether a paragraph holds any content. */
export const BOOKMARK_BOUNDARY_NODE_NAME = "bookmarkBoundary";

/**
 * Both boundary node types, because a bookmark pairs across them.
 *
 * A range that opens as a child of `w:body` and closes inside a paragraph is
 * the commoner shape in real documents, so a pass that read one node type
 * would see half of every such pair, call it unpaired, and delete it.
 */
const BOUNDARY_NODE_NAMES: ReadonlySet<string> = new Set([
  BOOKMARK_BOUNDARY_NODE_NAME,
  BLOCK_BOOKMARK_BOUNDARY_NODE_NAME,
]);

type PositionedNode = {
  position: number;
  node: PMNode;
};

type PositionedBoundary = BookmarkBoundaryOccurrence & PositionedNode;

/** The markers a `table` or `tableRow` node carries, or none. */
const positionedBookmarksOf = (node: PMNode): readonly PositionedBookmarkMarker[] => {
  const attrs = node.type.name === "table" ? readTableAttrs(node) : readTableRowAttrs(node);
  return attrs.ok ? (attrs.value._bookmarks ?? []) : [];
};

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
    // A marker a table or a row carries has no node of its own, so the pass
    // would call its partner unpaired and delete it. Counting it at the
    // container's position is what keeps a row-spanning bookmark whole while
    // its other half is being edited.
    if (node.type.name === "table" || node.type.name === "tableRow") {
      for (const { marker } of positionedBookmarksOf(node)) {
        boundaries.push({
          id: marker.id,
          type: marker.type === "bookmarkStart" ? "start" : "end",
          position,
          node,
        });
      }
    }
    if (!BOUNDARY_NODE_NAMES.has(node.type.name)) {
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
  // Only a boundary *node* can be deleted. A marker on a table or a row is an
  // attribute of a record the user is still editing, so an unpaired one is
  // reported by the pair's other half going, not by the container being cut.
  return [
    ...boundaries.filter(
      ({ id, node }) => invalidIds.has(id) && BOUNDARY_NODE_NAMES.has(node.type.name),
    ),
    ...malformedBoundaries,
  ];
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
    attrs: bookmarkBoundaryAttrSpec,
    parseDOM: [
      {
        tag: "span[data-docx-bookmark-boundary]",
        getAttrs: parseBookmarkBoundaryDom,
      },
    ],
    toDOM(node) {
      return ["span", bookmarkBoundaryDomAttributes(node, options.getInternalClipboardToken?.())];
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
