import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";

import type { DocxConformanceClass } from "../types/document";
import { getLocalName, getNamespacePrefix, parseXmlDocument } from "./xmlParser";

const STRICT_MAIN_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_MAIN_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export const detectDocxConformanceClass = (documentXml: string | null): DocxConformanceClass => {
  if (documentXml === null) {
    return DOCX_CONFORMANCE_CLASSES.UNKNOWN;
  }

  const root = parseXmlDocument(documentXml);
  if (root === null || getLocalName(root.name ?? "") !== "document") {
    return DOCX_CONFORMANCE_CLASSES.UNKNOWN;
  }

  const prefix = getNamespacePrefix(root.name ?? "");
  const namespaceAttribute = prefix === null ? "xmlns" : `xmlns:${prefix}`;
  const namespace = root.attributes?.[namespaceAttribute];
  if (namespace === STRICT_MAIN_NAMESPACE) {
    return DOCX_CONFORMANCE_CLASSES.STRICT;
  }
  if (namespace === TRANSITIONAL_MAIN_NAMESPACE) {
    return DOCX_CONFORMANCE_CLASSES.TRANSITIONAL;
  }
  return DOCX_CONFORMANCE_CLASSES.UNKNOWN;
};
