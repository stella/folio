/**
 * A bookmark boundary that stands between two blocks.
 *
 * The block-level twin of `bookmarkBoundary`. `CT_Body`, `CT_Tc` and
 * `CT_SdtContentBlock` declare `w:bookmarkStart`/`w:bookmarkEnd` beside their
 * paragraphs and tables, and Word writes them there whenever the selection was
 * whole blocks: `_GoBack` after an edit, a table-of-contents range around the
 * headings it lists. ProseMirror decides inline or block per node type, so the
 * inline atom cannot also sit here and this is its own node — with the same
 * attributes, read by the same reader, so the two levels cannot describe a
 * bookmark differently.
 *
 * **Position is structure, not an attribute**, for the reason
 * `PreservedBlockExtension` gives: the node sits between the same two blocks it
 * sat between in the source and ProseMirror's own mapping keeps it there, so an
 * edit inside the range cannot change what the range covers.
 *
 * Pairing is owned one level up. A bookmark may open at block level and close
 * inside a paragraph, which is the commoner shape in real documents, so the
 * integrity pass in `BookmarkBoundaryExtension` reads both node types together;
 * a rule enforced here would see half of each such pair and delete it.
 */

import { createNodeExtension } from "../create";

import {
  bookmarkBoundaryAttrSpec,
  bookmarkBoundaryDomAttributes,
  parseBookmarkBoundaryDom,
} from "./bookmarkBoundaryDom";

/** The node name, for callers asking what a block container holds. */
export const BLOCK_BOOKMARK_BOUNDARY_NODE_NAME = "blockBookmarkBoundary";

export const BlockBookmarkBoundaryExtension = createNodeExtension({
  name: BLOCK_BOOKMARK_BOUNDARY_NODE_NAME,
  schemaNodeName: BLOCK_BOOKMARK_BOUNDARY_NODE_NAME,
  nodeSpec: {
    group: "block",
    atom: true,
    selectable: false,
    attrs: bookmarkBoundaryAttrSpec,
    parseDOM: [
      {
        tag: "div[data-docx-bookmark-boundary]",
        getAttrs: parseBookmarkBoundaryDom,
      },
    ],
    toDOM(node) {
      return ["div", bookmarkBoundaryDomAttributes(node)];
    },
  },
});
