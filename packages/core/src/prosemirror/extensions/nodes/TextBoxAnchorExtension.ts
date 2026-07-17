/**
 * Zero-width inline anchor for a text box extracted to a sibling PM block.
 *
 * Keeping the anchor in the paragraph lets ordinary ProseMirror transactions
 * map its position as surrounding text is edited. The save conversion consumes
 * it when the sibling text box is restored to the DOCX run stream.
 */

import { expectTextBoxAnchorAttrs } from "../../textBoxAnchorAttrs";
import { createNodeExtension } from "../create";

export const TextBoxAnchorExtension = createNodeExtension({
  name: "textBoxAnchor",
  schemaNodeName: "textBoxAnchor",
  nodeSpec: {
    inline: true,
    group: "inline",
    marks: "_",
    atom: true,
    selectable: false,
    attrs: {
      anchorId: {},
    },
    parseDOM: [
      {
        tag: "span[data-docx-textbox-anchor]",
        getAttrs(dom): { anchorId: string } | false {
          if (!(dom instanceof HTMLElement)) {
            return false;
          }
          const anchorId = dom.dataset["docxTextboxAnchor"];
          return anchorId ? { anchorId } : false;
        },
      },
    ],
    toDOM(node) {
      const { anchorId } = expectTextBoxAnchorAttrs(node);
      return [
        "span",
        {
          "data-docx-textbox-anchor": anchorId,
          "aria-hidden": "true",
          contenteditable: "false",
          style:
            "display: inline-block; width: 0; height: 0; overflow: hidden; pointer-events: none;",
        },
      ];
    },
  },
});
