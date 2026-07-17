import { DOCX_CONFORMANCE_CLASSES } from "@stll/docx-core/model";

import type { DocxConformanceClass } from "../types/document";
import type { XmlElement } from "./xmlParser";
import { getLocalName, getNamespacePrefix, parseXmlDocument } from "./xmlParser";

const STRICT_MAIN_NAMESPACE = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_MAIN_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const ROOT_START_TAG_SCAN_LIMIT = 64 * 1024;
const XML_WHITESPACE = /\s/u;

/**
 * Find the root start tag within a bounded prefix of `document.xml`,
 * without backtracking.
 *
 * Equivalent in captured-tag semantics to the previous
 * `/^(?:\uFEFF|\s|<\?[\s\S]*?\?>|<!--[\s\S]*?-->)*(?<tag><[^\s/>]+(?:[^>"']|"[^"]*"|'[^']*')*>)/u`
 * pattern, but as a linear hand-rolled scanner: the original alternation of
 * greedy/lazy `[\s\S]*?` spans over an attacker-controlled prefix was
 * susceptible to catastrophic backtracking on crafted input. This walks
 * `scanTarget` once, tracking in-quote state, and returns `undefined` in
 * every case the regex would have failed to match (unterminated PI/comment,
 * no tag found before the scan limit, malformed tag open).
 */
function scanRootStartTag(scanTarget: string): string | undefined {
  const length = scanTarget.length;
  let index = 0;

  // Skip a BOM / whitespace / processing-instruction / comment prefix, in
  // any order and any number of times.
  for (;;) {
    if (index >= length) {
      return undefined;
    }
    const character = scanTarget[index];
    if (character === "\uFEFF" || XML_WHITESPACE.test(character ?? "")) {
      index += 1;
      continue;
    }
    if (scanTarget.startsWith("<?", index)) {
      const closeIndex = scanTarget.indexOf("?>", index + 2);
      if (closeIndex === -1) {
        return undefined;
      }
      index = closeIndex + 2;
      continue;
    }
    if (scanTarget.startsWith("<!--", index)) {
      const closeIndex = scanTarget.indexOf("-->", index + 4);
      if (closeIndex === -1) {
        return undefined;
      }
      index = closeIndex + 3;
      continue;
    }
    break;
  }

  // The root start tag must open here: `<` followed by at least one
  // character that is not whitespace, `/`, or `>`.
  if (scanTarget[index] !== "<") {
    return undefined;
  }
  const nameStart = index + 1;
  const firstNameChar = scanTarget[nameStart];
  if (
    firstNameChar === undefined ||
    firstNameChar === "/" ||
    firstNameChar === ">" ||
    XML_WHITESPACE.test(firstNameChar)
  ) {
    return undefined;
  }

  // Scan to the first unquoted `>`, honoring single- and double-quoted
  // attribute values so a `>` inside an attribute (e.g. `data-x="1 > 0"`)
  // does not end the tag early.
  let quote: '"' | "'" | undefined;
  for (let cursor = nameStart; cursor < length; cursor++) {
    const character = scanTarget[cursor];
    if (quote) {
      if (character === quote) {
        quote = undefined;
      }
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (character === ">") {
      return scanTarget.slice(index, cursor + 1);
    }
  }
  return undefined;
}

export const detectDocxConformanceClass = (documentXml: string | null): DocxConformanceClass => {
  if (documentXml === null) {
    return DOCX_CONFORMANCE_CLASSES.UNKNOWN;
  }

  const scanTarget =
    documentXml.length <= ROOT_START_TAG_SCAN_LIMIT
      ? documentXml
      : documentXml.slice(0, ROOT_START_TAG_SCAN_LIMIT);
  const rootStartTag = scanRootStartTag(scanTarget);
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
