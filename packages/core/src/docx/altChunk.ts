import {
  getLocalName,
  getChildElements,
  getAttribute,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";

export const ALT_CHUNK_READER_DIAGNOSTIC = "[Unsupported w:altChunk content]";
export const ALT_CHUNK_PLAIN_TEXT_LIMIT_BYTES = 1_000_000;
const OFFICE_RELATIONSHIPS_NAMESPACE_URI =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** The captured fragment's expanded name, not its source prefix, identifies w:altChunk. */
export const isAltChunkMarkup = (xml: string): boolean => {
  const element = getChildElements(parseXml(xml, OOXML_NAMESPACE_SCOPE)).at(0);
  return (
    element?.type === "element" &&
    getLocalName(element.name) === "altChunk" &&
    WORDPROCESSINGML_NAMESPACE_URIS.has(element.namespaceUri ?? "")
  );
};

export const getAltChunkRelationshipId = (xml: string): string | undefined => {
  const element = getChildElements(parseXml(xml, OOXML_NAMESPACE_SCOPE)).at(0);
  if (
    element?.type !== "element" ||
    getLocalName(element.name) !== "altChunk" ||
    !WORDPROCESSINGML_NAMESPACE_URIS.has(element.namespaceUri ?? "")
  ) {
    return undefined;
  }
  return getAttribute(element, OFFICE_RELATIONSHIPS_NAMESPACE_URI, "id") ?? undefined;
};
