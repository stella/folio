/** Hidden boundary for one non-empty tracked-move range. */

import { expectMoveRangeBoundaryAttrs, readMoveRangeMarker } from "../../moveRangeBoundaryAttrs";
import { createNodeExtension } from "../create";

type MoveRangeBoundaryOptions = {
  getInternalClipboardToken?: () => string;
};

export const MOVE_RANGE_BOUNDARY_NODE_NAME = "moveRangeBoundary";
export const MOVE_RANGE_BOUNDARY_DOM_ATTRIBUTE = "data-docx-move-range-boundary";

export const MoveRangeBoundaryExtension = createNodeExtension<MoveRangeBoundaryOptions>({
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
            const result = readMoveRangeMarker(marker);
            return result.ok ? { marker: result.value } : false;
          } catch {
            return false;
          }
        },
      },
    ],
    toDOM(node) {
      const marker = expectMoveRangeBoundaryAttrs(node);
      return [
        "span",
        {
          [MOVE_RANGE_BOUNDARY_DOM_ATTRIBUTE]: JSON.stringify(marker),
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
