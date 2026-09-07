/** A parser that routes its capture through the conversion owner. */

import { captureVerbatimXml } from "./verbatimCapture";
import { getLocalName, type XmlElement } from "./xmlParser";

export const captureProperties = (element: XmlElement): string =>
  getLocalName(element.name) === "tcPr" ? captureVerbatimXml(element) : "";
