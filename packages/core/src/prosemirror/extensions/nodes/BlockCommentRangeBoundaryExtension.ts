/**
 * A comment range boundary that stands between two blocks.
 *
 * The block-level twin of `commentRangeBoundary`, as `blockBookmarkBoundary`
 * is of `bookmarkBoundary`: `CT_Body`, `CT_Tc` and `CT_SdtContentBlock`
 * declare `w:commentRangeStart` and `w:commentRangeEnd` beside their blocks,
 * and a range that opens beside a content control is written there. The node
 * carries the same marker, read by the same reader, so the two levels cannot
 * describe a comment range differently. The comment's mark still covers the
 * text; the save keeps this node's marker as the range's own start or end.
 */

import {
  expectCommentRangeBoundaryAttrs,
  readCommentRangeBoundaryMarker,
} from "../../commentRangeBoundaryAttrs";
import { createNodeExtension } from "../create";
import type { NodeExtension } from "../types";

/** The node name, for callers asking what a block container holds. */
export const BLOCK_COMMENT_RANGE_BOUNDARY_NODE_NAME =
  "blockCommentRangeBoundary";
const BLOCK_COMMENT_RANGE_BOUNDARY_DOM_ATTRIBUTE =
  "data-docx-block-comment-range-boundary";

export const BlockCommentRangeBoundaryExtension: () => NodeExtension =
  createNodeExtension({
    name: BLOCK_COMMENT_RANGE_BOUNDARY_NODE_NAME,
    schemaNodeName: BLOCK_COMMENT_RANGE_BOUNDARY_NODE_NAME,
    nodeSpec: {
      group: "block",
      atom: true,
      selectable: false,
      attrs: { marker: {} },
      parseDOM: [
        {
          tag: `div[${BLOCK_COMMENT_RANGE_BOUNDARY_DOM_ATTRIBUTE}]`,
          getAttrs(dom) {
            const raw = dom.getAttribute(
              BLOCK_COMMENT_RANGE_BOUNDARY_DOM_ATTRIBUTE,
            );
            if (raw === null) return false;
            try {
              const marker = readCommentRangeBoundaryMarker(JSON.parse(raw));
              return marker ? { marker } : false;
            } catch {
              return false;
            }
          },
        },
      ],
      toDOM(node) {
        return [
          "div",
          {
            [BLOCK_COMMENT_RANGE_BOUNDARY_DOM_ATTRIBUTE]: JSON.stringify(
              expectCommentRangeBoundaryAttrs(node).marker,
            ),
            "aria-hidden": "true",
            contenteditable: "false",
            style: "display: none;",
          },
        ];
      },
    },
  });
