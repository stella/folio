/** Zero-width bookmark boundary that preserves its position through ProseMirror edits. */

import type { Node as PMNode } from "prosemirror-model";
import { Plugin } from "prosemirror-state";

import type { PositionedBookmarkMarker } from "../../../types/document";
import {
  readParagraphAttrs,
  readTableAttrs,
  readTableRowAttrs,
} from "../../attrs";
import { readBookmarkBoundaryAttrs } from "../../bookmarkBoundaryAttrs";
import {
  type BookmarkBoundaryOccurrence,
  findInvalidBookmarkBoundaryIds,
} from "../../bookmarkBoundaryIntegrity";
import { createNodeExtension } from "../create";
import type { NodeExtension } from "../types";
import { BLOCK_BOOKMARK_BOUNDARY_NODE_NAME } from "./BlockBookmarkBoundaryExtension";
import {
  bookmarkBoundaryAttrSpec,
  bookmarkBoundaryDomAttributes,
  parseBookmarkBoundaryDom,
} from "./bookmarkBoundaryDom";
import { INLINE_ANCHOR_DOM_STYLE } from "./inlineAnchorDom";

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
const positionedBookmarksOf = (
  node: PMNode,
): readonly PositionedBookmarkMarker[] => {
  const attrs =
    node.type.name === "table" ? readTableAttrs(node) : readTableRowAttrs(node);
  return attrs.ok ? (attrs.value._bookmarks ?? []) : [];
};

/** A boundary node to delete, by the id it names (none when malformed). */
type InvalidBoundary = PositionedNode & { id: number | null };

type BoundaryCensus = {
  /** Every boundary node of an id that does not pair, and every malformed one. */
  invalid: readonly InvalidBoundary[];
  /**
   * The ids that do not pair only because a partner is missing: at most one
   * start and one end, and no paragraph bookmark of the same id. A repeated
   * or colliding id is ambiguous, not lone, and is never kept.
   */
  unpairedIds: ReadonlySet<number>;
};

const censusByDoc = new WeakMap<PMNode, BoundaryCensus>();

/**
 * The census of a document, kept by the document: the state an edit starts
 * from is the one the previous edit produced, so each document is walked once.
 */
const boundaryCensus = (doc: PMNode): BoundaryCensus => {
  const cached = censusByDoc.get(doc);
  if (cached !== undefined) return cached;
  const census = collectInvalidBoundaries(doc);
  censusByDoc.set(doc, census);
  return census;
};

const collectInvalidBoundaries = (doc: PMNode): BoundaryCensus => {
  const boundaries: PositionedBoundary[] = [];
  const malformedBoundaries: InvalidBoundary[] = [];
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
      malformedBoundaries.push({ position, node, id: null });
      return false;
    }
    const attrs = result.value;
    const boundary = { id: attrs.id, type: attrs.type, position, node };
    boundaries.push(boundary);
    return false;
  });

  const invalidIds = findInvalidBookmarkBoundaryIds(
    boundaries,
    paragraphBookmarkIds,
  );
  const counts = new Map<number, { starts: number; ends: number }>();
  for (const { id, type } of boundaries) {
    const count = counts.get(id) ?? { starts: 0, ends: 0 };
    if (type === "start") count.starts += 1;
    else count.ends += 1;
    counts.set(id, count);
  }
  const unpairedIds = new Set<number>();
  for (const id of invalidIds) {
    const count = counts.get(id);
    if (
      count !== undefined &&
      count.starts <= 1 &&
      count.ends <= 1 &&
      !paragraphBookmarkIds.has(id)
    )
      unpairedIds.add(id);
  }
  // Only a boundary *node* can be deleted. A marker on a table or a row is an
  // attribute of a record the user is still editing, so an unpaired one is
  // reported by the pair's other half going, not by the container being cut.
  return {
    invalid: [
      ...boundaries.filter(
        ({ id, node }) =>
          invalidIds.has(id) && BOUNDARY_NODE_NAMES.has(node.type.name),
      ),
      ...malformedBoundaries,
    ],
    unpairedIds,
  };
};

export const BookmarkBoundaryExtension: (
  options?: Partial<BookmarkBoundaryOptions> | undefined,
) => NodeExtension = createNodeExtension<BookmarkBoundaryOptions>({
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
        getAttrs: (dom) =>
          parseBookmarkBoundaryDom(dom, options.getInternalClipboardToken?.()),
      },
    ],
    toDOM(node) {
      return [
        "span",
        {
          ...bookmarkBoundaryDomAttributes(
            node,
            options.getInternalClipboardToken?.(),
          ),
          style: INLINE_ANCHOR_DOM_STYLE,
        },
      ];
    },
  }),
  onSchemaReady: () => ({
    plugins: [
      new Plugin({
        appendTransaction(transactions, oldState, newState) {
          if (!transactions.some(({ docChanged }) => docChanged)) {
            return null;
          }

          const { invalid, unpairedIds } = boundaryCensus(newState.doc);
          if (invalid.length === 0) {
            return null;
          }
          // A bookmark the document already held without a partner before
          // these edits is source content: a lone `_GoBack` start, or a start
          // whose end the source never had. The edits did not break it, so
          // the pass leaves it as it was read. Only a pair the edits broke,
          // or a boundary they brought in alone, is removed.
          const { unpairedIds: alreadyUnpaired } = boundaryCensus(oldState.doc);
          const invalidBoundaries = invalid.filter(
            ({ id }) =>
              id === null || !unpairedIds.has(id) || !alreadyUnpaired.has(id),
          );
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
