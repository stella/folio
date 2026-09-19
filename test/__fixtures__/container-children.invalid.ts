/**
 * A container walked by hand: the cases say what is modelled and the `default`
 * says nothing, so every child the schema adds next is dropped in silence.
 * The lint rule rejects this shape outside the dispatcher.
 */

import {
  getChildElements,
  getLocalName,
  type XmlElement,
} from "../../packages/core/src/docx/xmlParser";

export const parseThing = (container: XmlElement): string[] => {
  const modelled: string[] = [];
  for (const child of getChildElements(container)) {
    const localName = getLocalName(child.name);
    switch (localName) {
      case "p":
        modelled.push("paragraph");
        break;
      case "tbl":
        modelled.push("table");
        break;
      case "sdt":
        modelled.push("control");
        break;
      default:
        break;
    }
  }
  return modelled;
};
