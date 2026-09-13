import { captureVerbatimXml } from "./verbatimCapture";

import {
  findAttributeByNamespaceUri,
  getLocalName,
  getNamespaceUri,
  OOXML_NAMESPACE_SCOPE,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  parseXml,
  type XmlElement,
} from "./xmlParser";

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
      const attribute = findAttributeByNamespaceUri(
        element,
        OFFICE_RELATIONSHIP_NAMESPACE_URIS,
        localName,
      );
      if (!attribute || attribute.name !== name) continue;
      const imageAttribute =
        localName === "embed" ||
        (localName === "id" &&
          getLocalName(element.name) === "imagedata" &&
          getNamespaceUri(element) === "urn:schemas-microsoft-com:vml");
      if (!imageAttribute || attribute.value !== previousId) {
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
