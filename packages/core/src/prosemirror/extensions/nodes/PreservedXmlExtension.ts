/**
 * An opaque inline atom for a run child folio does not model.
 *
 * The editor cannot edit markup it has no model for, and it does not have to:
 * it only has to not lose it. The node carries the captured bytes untouched
 * and shows the visible text the markup puts on the line (`w:ruby` renders its
 * `w:rubyBase` there), so the round trip keeps both the document's text and
 * the document's markup.
 */

import { expectPreservedXmlAttrs } from "../../attrs";
import { PRESERVED_XML_LEVELS } from "../../schema/nodes";
import { createNodeExtension } from "../create";

/** The node name, for callers asking what a paragraph holds. */
export const PRESERVED_XML_NODE_NAME = "preservedXml";

export const PreservedXmlExtension = createNodeExtension({
  name: PRESERVED_XML_NODE_NAME,
  schemaNodeName: PRESERVED_XML_NODE_NAME,
  nodeSpec: {
    inline: true,
    group: "inline",
    atom: true,
    marks: "_",
    selectable: false,
    attrs: {
      xml: {},
      text: { default: "" },
      // A run child is the case a paste through the DOM is most likely to
      // carry, and the one whose re-wrapping in a `w:r` is always legal.
      level: { default: PRESERVED_XML_LEVELS.run },
      // `undefined`, so a capture that is not a folded list-number field
      // serializes exactly as it did before the attr existed.
      foldedListNumber: { default: undefined },
    },
    parseDOM: [
      {
        tag: "span[data-docx-preserved-xml]",
        getAttrs(node) {
          if (!(node instanceof HTMLElement)) {
            return false;
          }
          const xml = node.dataset["docxPreservedXml"];
          if (xml === undefined || xml === "") {
            return false;
          }
          const level =
            node.dataset["docxPreservedLevel"] === PRESERVED_XML_LEVELS.inline
              ? PRESERVED_XML_LEVELS.inline
              : PRESERVED_XML_LEVELS.run;
          const folded = node.dataset["docxFoldedListNumber"];
          return {
            xml,
            text: node.dataset["docxPreservedText"] ?? "",
            level,
            ...(folded === "field" || folded === "tab" ? { foldedListNumber: folded } : {}),
          };
        },
      },
    ],
    toDOM(node) {
      const { xml, text, level, foldedListNumber } = expectPreservedXmlAttrs(node);
      return [
        "span",
        {
          "data-docx-preserved-xml": xml,
          "data-docx-preserved-level": level,
          ...(text === "" ? {} : { "data-docx-preserved-text": text }),
          ...(foldedListNumber ? { "data-docx-folded-list-number": foldedListNumber } : {}),
          contenteditable: "false",
        },
        text,
      ];
    },
  },
});
