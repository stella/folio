import {
  expectCommentRangeBoundaryAttrs,
  readCommentRangeBoundaryMarker,
} from "../../commentRangeBoundaryAttrs";
import { createNodeExtension } from "../create";
import type { NodeExtension } from "../types";

type CommentRangeBoundaryOptions = {
  getInternalClipboardToken?: () => string;
};

export const COMMENT_RANGE_BOUNDARY_NODE_NAME = "commentRangeBoundary";
export const COMMENT_RANGE_BOUNDARY_DOM_ATTRIBUTE =
  "data-docx-comment-range-boundary";

export const CommentRangeBoundaryExtension: (
  options?: Partial<CommentRangeBoundaryOptions> | undefined,
) => NodeExtension = createNodeExtension<CommentRangeBoundaryOptions>({
  name: COMMENT_RANGE_BOUNDARY_NODE_NAME,
  schemaNodeName: COMMENT_RANGE_BOUNDARY_NODE_NAME,
  nodeSpec: (options) => ({
    inline: true,
    group: "inline",
    marks: "_",
    atom: true,
    selectable: false,
    attrs: { marker: {} },
    parseDOM: [
      {
        tag: `span[${COMMENT_RANGE_BOUNDARY_DOM_ATTRIBUTE}]`,
        getAttrs(dom) {
          const raw = dom.getAttribute(COMMENT_RANGE_BOUNDARY_DOM_ATTRIBUTE);
          if (raw === null) return false;
          try {
            const marker: unknown = JSON.parse(raw);
            const value = readCommentRangeBoundaryMarker(marker);
            return value ? { marker: value } : false;
          } catch {
            return false;
          }
        },
      },
    ],
    toDOM(node) {
      const { marker } = expectCommentRangeBoundaryAttrs(node);
      return [
        "span",
        {
          [COMMENT_RANGE_BOUNDARY_DOM_ATTRIBUTE]: JSON.stringify(marker),
          "aria-hidden": "true",
          contenteditable: "false",
          style: "display: none;",
          ...(options.getInternalClipboardToken
            ? {
                "data-docx-internal-clipboard":
                  options.getInternalClipboardToken(),
              }
            : {}),
        },
      ];
    },
  }),
});
