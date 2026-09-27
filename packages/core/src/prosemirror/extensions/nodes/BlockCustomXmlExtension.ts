import { createNodeExtension } from "../create";

/** An editable block sequence inside an authored custom XML wrapper. */
export const BlockCustomXmlExtension = createNodeExtension({
  name: "blockCustomXml",
  schemaNodeName: "blockCustomXml",
  nodeSpec: {
    group: "block",
    content: "block+",
    isolating: true,
    defining: true,
    attrs: {
      openingXml: { default: null },
      closingXml: { default: null },
      _originallyEmpty: { default: false },
    },
    parseDOM: [
      {
        tag: "div[data-docx-block-custom-xml-open]",
        getAttrs(dom) {
          if (!(dom instanceof HTMLElement)) {
            return false;
          }
          const openingXml = dom.dataset["docxBlockCustomXmlOpen"];
          const closingXml = dom.dataset["docxBlockCustomXmlClose"];
          if (!openingXml || !closingXml) {
            return false;
          }
          return {
            openingXml,
            closingXml,
            _originallyEmpty: dom.dataset["docxBlockCustomXmlEmpty"] === "true",
          };
        },
      },
    ],
    toDOM(node) {
      const { openingXml, closingXml, _originallyEmpty } = node.attrs;
      if (typeof openingXml !== "string" || typeof closingXml !== "string") {
        return ["div", { class: "docx-block-custom-xml" }, 0];
      }
      return [
        "div",
        {
          class: "docx-block-custom-xml",
          "data-docx-block-custom-xml-open": openingXml,
          "data-docx-block-custom-xml-close": closingXml,
          "data-docx-block-custom-xml-empty": String(_originallyEmpty === true),
        },
        0,
      ];
    },
  },
});
