/** Zero-width cached pagination boundary preserved from DOCX. */

import { createNodeExtension } from "../create";
import type { NodeExtension } from "../types";

/**
 * The node name, for callers that have to ask whether anything is LEFT in a
 * paragraph. This node is zero-width and carries no revision of its own, so a
 * paragraph holding only these holds nothing a reader would see.
 */
export const RENDERED_PAGE_BREAK_NODE_NAME = "renderedPageBreak";

export const RenderedPageBreakExtension: (
  options?: Partial<Record<string, unknown>> | undefined,
) => NodeExtension = createNodeExtension({
  name: RENDERED_PAGE_BREAK_NODE_NAME,
  schemaNodeName: RENDERED_PAGE_BREAK_NODE_NAME,
  nodeSpec: {
    inline: true,
    group: "inline",
    atom: true,
    selectable: false,
    parseDOM: [{ tag: "span[data-docx-rendered-page-break]" }],
    toDOM() {
      return [
        "span",
        {
          "data-docx-rendered-page-break": "true",
          style: "display: none;",
        },
      ];
    },
  },
});
