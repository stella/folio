/**
 * A tracked-move range boundary that stands between two blocks.
 *
 * The block-level twin of `moveRangeBoundary`, as `blockBookmarkBoundary` is
 * of `bookmarkBoundary`: `CT_Body`, `CT_Tc` and `CT_SdtContentBlock` declare
 * `w:moveFromRangeStart`/`End` and `w:moveToRangeStart`/`End` beside their
 * paragraphs and tables, and a move whose range opens or closes on whole
 * blocks is written there. The node carries the same marker, read by the same
 * reader, so the two levels cannot describe a move range differently.
 *
 * Position is structure, not an attribute: the node sits between the same two
 * blocks it sat between in the source, and ProseMirror's mapping keeps it
 * there.
 */

import {
  INTERNAL_CLIPBOARD_TOKEN_ATTRIBUTE,
  isInternalClipboardDom,
} from "../../copyOriginDom";
import {
  expectMoveRangeBoundaryAttrs,
  moveRangeMarkerForDom,
  readMoveRangeMarker,
  withoutUntrustedCopyOrigin,
} from "../../moveRangeBoundaryAttrs";
import { createNodeExtension } from "../create";

type BlockMoveRangeBoundaryOptions = {
  getInternalClipboardToken?: () => string;
};

/** The node name, for callers asking what a block container holds. */
export const BLOCK_MOVE_RANGE_BOUNDARY_NODE_NAME = "blockMoveRangeBoundary";
const BLOCK_MOVE_RANGE_BOUNDARY_DOM_ATTRIBUTE =
  "data-docx-block-move-range-boundary";

export const BlockMoveRangeBoundaryExtension =
  createNodeExtension<BlockMoveRangeBoundaryOptions>({
    name: BLOCK_MOVE_RANGE_BOUNDARY_NODE_NAME,
    schemaNodeName: BLOCK_MOVE_RANGE_BOUNDARY_NODE_NAME,
    nodeSpec: (options) => ({
      group: "block",
      atom: true,
      selectable: false,
      attrs: { marker: {} },
      parseDOM: [
        {
          tag: `div[${BLOCK_MOVE_RANGE_BOUNDARY_DOM_ATTRIBUTE}]`,
          getAttrs(dom) {
            const raw = dom.getAttribute(
              BLOCK_MOVE_RANGE_BOUNDARY_DOM_ATTRIBUTE,
            );
            if (raw === null) {
              return false;
            }
            try {
              const marker: unknown = JSON.parse(raw);
              const markerForRead = isInternalClipboardDom(
                dom,
                options.getInternalClipboardToken?.(),
              )
                ? marker
                : withoutUntrustedCopyOrigin(marker);
              const result = readMoveRangeMarker(markerForRead);
              return result.ok ? { marker: result.value } : false;
            } catch {
              return false;
            }
          },
        },
      ],
      toDOM(node) {
        const internalClipboardToken = options.getInternalClipboardToken?.();
        const marker = moveRangeMarkerForDom(
          expectMoveRangeBoundaryAttrs(node),
          internalClipboardToken,
        );
        return [
          "div",
          {
            [BLOCK_MOVE_RANGE_BOUNDARY_DOM_ATTRIBUTE]: JSON.stringify(marker),
            "aria-hidden": "true",
            contenteditable: "false",
            style: "display: none;",
            ...(internalClipboardToken
              ? {
                  [INTERNAL_CLIPBOARD_TOKEN_ATTRIBUTE]: internalClipboardToken,
                }
              : {}),
          },
        ];
      },
    }),
  });
