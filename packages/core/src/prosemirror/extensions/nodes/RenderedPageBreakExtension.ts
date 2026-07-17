/** Zero-width cached pagination boundary preserved from DOCX. */

import { createNodeExtension } from "../create";

export const RenderedPageBreakExtension = createNodeExtension({
  name: "renderedPageBreak",
  schemaNodeName: "renderedPageBreak",
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
