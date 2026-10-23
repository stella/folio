import type { Node as ProseMirrorNode } from "prosemirror-model";

import { decodeOoxmlSymbolCharacter } from "../utils/ooxmlSymbol";
import {
  expectFieldAttrs,
  expectPreservedXmlAttrs,
  expectSymbolAttrs,
} from "./attrs";

/** Native note-reference labels are navigation anchors, outside searchable text. */
export const isNoteReferenceText = (node: ProseMirrorNode): boolean =>
  node.isText && node.marks.some((mark) => mark.type.name === "footnoteRef");

/** A leaf's logical text; null means a transparent container to descend into. */
export const getInlineDisplayToken = (node: ProseMirrorNode): string | null => {
  if (isNoteReferenceText(node)) return "";
  if (node.type.name === "field") return expectFieldAttrs(node).displayText;
  if (node.type.name === "preservedXml")
    return expectPreservedXmlAttrs(node).text;
  if (node.isText) return node.text ?? "";
  if (node.type.name === "tab") return "\t";
  if (node.type.name === "hardBreak") return "\n";
  if (node.type.name === "symbol")
    return decodeOoxmlSymbolCharacter(expectSymbolAttrs(node).char) ?? "";
  return null;
};

/** Visible inline text, including cached source atoms and native run tokens. */
export const getInlineDisplayText = (paragraph: ProseMirrorNode): string => {
  let text = "";
  paragraph.descendants((node) => {
    const token = getInlineDisplayToken(node);
    if (token === null) return true;
    text += token;
    return false;
  });
  return text;
};
