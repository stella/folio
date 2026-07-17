import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";

import type { DocxConformanceClass } from "../types/document";
import type { XmlElement } from "./xmlParser";
import { getLocalName, getNamespacePrefix, parseXmlDocument } from "./xmlParser";

const STRICT_MAIN_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_MAIN_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const ROOT_START_TAG_SCAN_LIMIT = 64 * 1024;
const ROOT_START_TAG_PATTERN =
  /^(?:\uFEFF|\s|<\?[\s\S]*?\?>|<!--[\s\S]*?-->)*(?<tag><[^\s/>]+(?:[^>"']|"[^"]*"|'[^']*')*>)/u;

export const detectDocxConformanceClass = (documentXml: string | null): DocxConformanceClass => {
  if (documentXml === null) {
    return DOCX_CONFORMANCE_CLASSES.UNKNOWN;
  }

  const scanTarget =
    documentXml.length <= ROOT_START_TAG_SCAN_LIMIT
      ? documentXml
      : documentXml.slice(0, ROOT_START_TAG_SCAN_LIMIT);
  const rootStartTag = ROOT_START_TAG_PATTERN.exec(scanTarget)?.groups?.["tag"];
  const fastRoot =
    rootStartTag === undefined
      ? null
      : parseXmlDocument(
          rootStartTag.endsWith("/>") ? rootStartTag : `${rootStartTag.slice(0, -1)}/>`,
        );
  // Preserve the full parser's behavior for uncommon XML prologs that the
  // bounded fast path intentionally does not recognize (for example a DTD).
  const root = fastRoot ?? parseXmlDocument(documentXml);
  return detectDocxConformanceClassFromRoot(root);
};

const detectDocxConformanceClassFromRoot = (root: XmlElement | null): DocxConformanceClass => {
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
