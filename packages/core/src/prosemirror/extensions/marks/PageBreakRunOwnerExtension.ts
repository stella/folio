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

const CANONICAL_OWNER_ID = /^(?:0|[1-9]\d*)$/u;

export const PageBreakRunOwnerExtension = createMarkExtension({
  name: "pageBreakRunOwner",
  schemaMarkName: "pageBreakRunOwner",
  markSpec: {
    attrs: { id: {} },
    inclusive: false,
    parseDOM: [
      {
        tag: "span[data-docx-page-break-run-owner]",
        getAttrs(dom) {
          const rawId = dom.dataset["docxPageBreakRunOwner"];
          if (rawId === undefined || !CANONICAL_OWNER_ID.test(rawId)) {
            return false;
          }
          const id = Number(rawId);
          return Number.isSafeInteger(id) ? { id } : false;
        },
      },
    ],
    toDOM(mark) {
      const { id } = expectPageBreakRunOwnerMarkAttrs(mark);
      return ["span", { "data-docx-page-break-run-owner": String(id) }, 0];
    },
  },
});
