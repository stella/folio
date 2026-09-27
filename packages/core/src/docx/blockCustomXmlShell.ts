import type { BlockCustomXml } from "../types/document";
import { captureVerbatimXml } from "./verbatimCapture";
import { cloneElement, type XmlElement } from "./xmlParser";

/** Capture the wrapper independently of its modeled and preserved children. */
export const blockCustomXmlShell = (
  element: XmlElement,
): Pick<BlockCustomXml, "openingXml" | "closingXml"> => {
  const shell = captureVerbatimXml(cloneElement(element, { elements: [] }));
  const name = element.name ?? "w:customXml";
  return {
    openingXml: shell.endsWith("/>") ? `${shell.slice(0, -2)}>` : shell,
    closingXml: `</${name}>`,
  };
};
