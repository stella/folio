/** Inline owner for an authored `<w:br w:type="page"/>` run child. */

import { expectPageBreakRunAttrs } from "../../attrs";
import { createNodeExtension } from "../create";

const PAGE_BREAK_RUN_NODE_NAME = "pageBreakRun";

export const PageBreakRunExtension = createNodeExtension({
  name: PAGE_BREAK_RUN_NODE_NAME,
  schemaNodeName: PAGE_BREAK_RUN_NODE_NAME,
  nodeSpec: {
    inline: true,
    group: "inline",
    atom: true,
    marks: "_",
    attrs: {
      clear: { default: null },
    },
    selectable: true,
    parseDOM: [
      {
        tag: "span[data-docx-page-break-run]",
        getAttrs(node) {
          if (!(node instanceof HTMLElement)) {
            return false;
          }
          const clear = node.dataset["docxBreakClear"];
          return clear === "none" || clear === "left" || clear === "right" || clear === "all"
            ? { clear }
            : null;
        },
      },
    ],
    toDOM(node) {
      const { clear } = expectPageBreakRunAttrs(node);
      return [
        "span",
        {
          "aria-hidden": "true",
          "data-docx-page-break-run": "true",
          ...(clear ? { "data-docx-break-clear": clear } : {}),
          style: "display: inline-block; overflow: hidden; width: 0;",
        },
      ];
    },
  },
});
