import { captureVerbatimXml } from "./verbatimCapture";
import { OOXML_NS } from "@stll/docx-utils";

import {
  findAttributeByNamespaceUri,
  getLocalName,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  type XmlElement,
} from "./xmlParser";

const RELATIONSHIP_NAMESPACES: ReadonlySet<string> = new Set([
  OOXML_NS.r,
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

type RebindDrawingImageRelationshipOptions = {
  xml: string;
  previousId: string;
  nextId: string;
};

/** Only a single embedded image can travel without importing other package resources. */
export const rebindDrawingImageRelationship = ({
  xml,
  previousId,
  nextId,
}: RebindDrawingImageRelationshipOptions): string | null => {
  const root = parseXml(xml, OOXML_NAMESPACE_SCOPE);
  let embeddedCount = 0;
  let unsupported = false;
  const visit = (element: XmlElement): void => {
    for (const name of Object.keys(element.attributes ?? {})) {
      const localName = getLocalName(name);
      const attribute = findAttributeByNamespaceUri(element, RELATIONSHIP_NAMESPACES, localName);
      if (!attribute || attribute.name !== name) continue;
      if (localName !== "embed" || attribute.value !== previousId) {
        unsupported = true;
        continue;
      }
      embeddedCount++;
      if (element.attributes) element.attributes[name] = nextId;
    }
    for (const child of element.elements ?? []) visit(child);
  };
  visit(root);
  if (unsupported || embeddedCount !== 1) return null;
  return (root.elements ?? []).map(captureVerbatimXml).join("");
};
