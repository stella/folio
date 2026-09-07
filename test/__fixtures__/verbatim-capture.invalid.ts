/** Deliberate violation: a parser that serializes a capture itself. */

import { elementToXml, type XmlElement } from "./xmlParser";

export const captureProperties = (element: XmlElement): string => elementToXml(element);
