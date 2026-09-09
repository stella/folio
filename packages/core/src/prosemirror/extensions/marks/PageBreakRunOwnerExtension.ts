/**
 * Editor-only identity for the authored run surrounding a page-break atom.
 *
 * Ordinary formatting marks cannot distinguish two adjacent source runs with
 * identical properties. This mark is added only to a run containing an
 * explicit page break, so every sibling carrier can be rejoined exactly on
 * save without changing the generic run-coalescing model.
 */

import { expectPageBreakRunOwnerMarkAttrs } from "../../attrs";
import { createMarkExtension } from "../create";

export const PageBreakRunOwnerExtension = createMarkExtension({
  name: "pageBreakRunOwner",
  schemaMarkName: "pageBreakRunOwner",
  markSpec: {
    attrs: { id: { default: 0 } },
    inclusive: false,
    parseDOM: [
      {
        tag: "span[data-docx-page-break-run-owner]",
        getAttrs(dom) {
          const id = Number.parseInt(dom.dataset["docxPageBreakRunOwner"] ?? "", 10);
          return Number.isInteger(id) && id >= 0 ? { id } : false;
        },
      },
    ],
    toDOM(mark) {
      const { id } = expectPageBreakRunOwnerMarkAttrs(mark);
      return ["span", { "data-docx-page-break-run-owner": String(id) }, 0];
    },
  },
});
