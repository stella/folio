import type { Node as ProseMirrorNode } from "prosemirror-model";
import { decodeOoxmlSymbolCharacter } from "../utils/ooxmlSymbol";
import { expectFieldAttrs, expectPreservedXmlAttrs, expectSymbolAttrs } from "./attrs";

/** Visible inline text, including cached source atoms and native run tokens. */
export const getInlineDisplayText = (paragraph: ProseMirrorNode): string => {
  let text = "";
  paragraph.descendants((node) => {
    if (node.type.name === "field") {
      text += expectFieldAttrs(node).displayText;
      return false;
    }
    if (node.type.name === "preservedXml") {
      text += expectPreservedXmlAttrs(node).text;
      return false;
    }
    if (node.isText) {
      text += node.text ?? "";
      return false;
    }
    if (node.type.name === "tab") {
      text += "\t";
      return false;
    }
    if (node.type.name === "hardBreak") {
      text += "\n";
      return false;
    }
    if (node.type.name === "symbol") {
      text += decodeOoxmlSymbolCharacter(expectSymbolAttrs(node).char) ?? "";
      return false;
    }
    return true;
  });
  return text;
};
