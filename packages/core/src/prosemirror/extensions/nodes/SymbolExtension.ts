import { isOoxmlSymbolCharacter } from "@stll/docx-core/model";

import { expectSymbolAttrs } from "../../attrs";
import { decodeOoxmlSymbolCharacter } from "../../../utils/ooxmlSymbol";
import { createNodeExtension } from "../create";

const escapeCssString = (value: string): string =>
  value
    .replaceAll("\\", "\\\\")
    .replaceAll('"', '\\"')
    .replaceAll("\n", "\\a ")
    .replaceAll("\r", "\\d ")
    .replaceAll("\f", "\\c ");

export const SymbolExtension = createNodeExtension({
  name: "symbol",
  schemaNodeName: "symbol",
  nodeSpec: {
    inline: true,
    group: "inline",
    atom: true,
    marks: "_",
    attrs: {
      font: {},
      char: {},
    },
    selectable: false,
    parseDOM: [
      {
        tag: "span[data-docx-symbol-char][data-docx-symbol-font]",
        getAttrs(node) {
          if (!(node instanceof HTMLElement)) {
            return false;
          }
          const font = node.dataset["docxSymbolFont"];
          const char = node.dataset["docxSymbolChar"];
          if (!font || font.trim() === "" || !char || !isOoxmlSymbolCharacter(char)) {
            return false;
          }
          return { font, char };
        },
      },
    ],
    toDOM(node) {
      const { font, char } = expectSymbolAttrs(node);
      const decoded = decodeOoxmlSymbolCharacter(char);
      return [
        "span",
        {
          "data-docx-symbol-char": char,
          "data-docx-symbol-font": font,
          style: `font-family: "${escapeCssString(font)}";`,
        },
        decoded ?? "\uFFFD",
      ];
    },
  },
});
