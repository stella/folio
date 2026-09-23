import type { HeaderFooterType } from "../types/document";
import { parseHeaderFooterType } from "./headerFooterRefParser";
import {
  getAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  getNamespaceUri,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  parseXml,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";
import { assertXmlResourceLimits } from "./xmlResourceLimits";

export type HeaderFooterReference = {
  element: "headerReference" | "footerReference";
  /** The parsed `ST_HdrFtr` value, never the raw attribute. */
  type: HeaderFooterType;
  rId: string;
};

/** What a save must not lose from `word/document.xml`. */
export type DocumentSectionFacts = {
  /** `w:sectPr` elements, excluding the prior records inside `w:sectPrChange`. */
  sectionCount: number;
  /** Every header/footer reference, including those inside `w:sectPrChange`. */
  headerFooterReferences: readonly HeaderFooterReference[];
};

const isWordprocessingElement = (element: XmlElement, localName: string): boolean =>
  getLocalName(element.name) === localName &&
  WORDPROCESSINGML_NAMESPACE_URIS.has(getNamespaceUri(element) ?? "");

/** Both facts from one walk of a parsed document tree. */
export const collectDocumentSectionFacts = (root: XmlElement): DocumentSectionFacts => {
  let sectionCount = 0;
  const headerFooterReferences: HeaderFooterReference[] = [];
  const pending: Array<{ element: XmlElement; insidePropertyChange: boolean }> = [
    { element: root, insidePropertyChange: false },
  ];
  for (let current = pending.pop(); current !== undefined; current = pending.pop()) {
    const { element, insidePropertyChange } = current;
    const localName = getLocalName(element.name);
    if (!insidePropertyChange && isWordprocessingElement(element, "sectPr")) {
      sectionCount += 1;
    }
    if (
      (localName === "headerReference" || localName === "footerReference") &&
      isWordprocessingElement(element, localName)
    ) {
      const rId = getAttributeByNamespaceUri(element, OFFICE_RELATIONSHIP_NAMESPACE_URIS, "id");
      if (rId) {
        headerFooterReferences.push({
          element: localName,
          // Read through the parser: a `w:type` outside `ST_HdrFtr` that the
          // parse boundary normalised would otherwise read as a dropped
          // reference when the serialized package states the normalised value.
          type: parseHeaderFooterType(
            getAttributeByNamespaceUri(element, WORDPROCESSINGML_NAMESPACE_URIS, "type"),
          ),
          rId,
        });
      }
    }
    const childrenInsidePropertyChange =
      insidePropertyChange || isWordprocessingElement(element, "sectPrChange");
    for (const child of getChildElements(element)) {
      pending.push({ element: child, insidePropertyChange: childrenInsidePropertyChange });
    }
  }
  return { sectionCount, headerFooterReferences };
};

/** Read the section facts of a `word/document.xml` string under the shared resource limits. */
export const readDocumentSectionFacts = (xml: string): DocumentSectionFacts => {
  assertXmlResourceLimits({ xml });
  return collectDocumentSectionFacts(parseXml(xml));
};
