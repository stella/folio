import {
  getChildElements,
  getLocalName,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  WORDPROCESSINGML_NAMESPACE_URIS,
} from "./xmlParser";

export const OPAQUE_REVISION_CARRIER_READER_DIAGNOSTIC =
  "[Unsupported block-level tracked-change content]";

const REVISION_WRAPPERS: ReadonlySet<string> = new Set(["ins", "del", "moveFrom", "moveTo"]);

export const opaqueRevisionCarrierName = (xml: string): string | undefined => {
  const visit = (element: ReturnType<typeof getChildElements>[number]): string | undefined => {
    const localName = getLocalName(element.name);
    if (
      WORDPROCESSINGML_NAMESPACE_URIS.has(element.namespaceUri ?? "") &&
      REVISION_WRAPPERS.has(localName)
    ) {
      return `w:${localName}`;
    }
    for (const child of getChildElements(element)) {
      const match = visit(child);
      if (match !== undefined) {
        return match;
      }
    }
    return undefined;
  };
  for (const root of getChildElements(parseXml(xml, OOXML_NAMESPACE_SCOPE))) {
    const match = visit(root);
    if (match !== undefined) {
      return match;
    }
  }
  return undefined;
};

export const countOpaqueRevisionWrappers = (xml: string): number => {
  const count = (element: ReturnType<typeof getChildElements>[number]): number => {
    const localName = getLocalName(element.name);
    let total =
      WORDPROCESSINGML_NAMESPACE_URIS.has(element.namespaceUri ?? "") &&
      REVISION_WRAPPERS.has(localName)
        ? 1
        : 0;
    for (const child of getChildElements(element)) {
      total += count(child);
    }
    return total;
  };
  return getChildElements(parseXml(xml, OOXML_NAMESPACE_SCOPE)).reduce(
    (total, element) => total + count(element),
    0,
  );
};

export const isOpaqueNestedRowMarkup = (xml: string): boolean => {
  const element = getChildElements(parseXml(xml, OOXML_NAMESPACE_SCOPE)).at(0);
  return (
    element !== undefined &&
    getLocalName(element.name) === "tr" &&
    WORDPROCESSINGML_NAMESPACE_URIS.has(element.namespaceUri ?? "")
  );
};
