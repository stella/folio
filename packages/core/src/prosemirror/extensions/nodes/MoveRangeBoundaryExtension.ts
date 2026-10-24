/** Hidden boundary for one non-empty tracked-move range. */

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

type MoveRangeBoundaryOptions = {
  getInternalClipboardToken?: () => string;
};

export const MOVE_RANGE_BOUNDARY_NODE_NAME = "moveRangeBoundary";
export const MOVE_RANGE_BOUNDARY_DOM_ATTRIBUTE =
  "data-docx-move-range-boundary";

export const MoveRangeBoundaryExtension =
  createNodeExtension<MoveRangeBoundaryOptions>({
    name: MOVE_RANGE_BOUNDARY_NODE_NAME,
    schemaNodeName: MOVE_RANGE_BOUNDARY_NODE_NAME,
    nodeSpec: (options) => ({
      inline: true,
      group: "inline",
      marks: "_",
      atom: true,
      selectable: false,
      attrs: { marker: {} },
      parseDOM: [
        {
          tag: `span[${MOVE_RANGE_BOUNDARY_DOM_ATTRIBUTE}]`,
          getAttrs(dom) {
            const raw = dom.getAttribute(MOVE_RANGE_BOUNDARY_DOM_ATTRIBUTE);
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
          "span",
          {
            [MOVE_RANGE_BOUNDARY_DOM_ATTRIBUTE]: JSON.stringify(marker),
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
