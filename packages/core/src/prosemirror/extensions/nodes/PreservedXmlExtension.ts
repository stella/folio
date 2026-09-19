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
          return { xml, text: node.dataset["docxPreservedText"] ?? "" };
        },
      },
    ],
    toDOM(node) {
      const { xml, text } = expectPreservedXmlAttrs(node);
      return [
        "span",
        {
          "data-docx-preserved-xml": xml,
          ...(text === "" ? {} : { "data-docx-preserved-text": text }),
          contenteditable: "false",
        },
        text,
      ];
    },
  },
});
