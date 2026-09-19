/**
 * Header/Footer Reference Parser
 *
 * Parses header/footer references (w:headerReference, w:footerReference) that
 * appear in section properties. Extracted from headerFooterParser to break the
 * circular dependency: headerFooterParser -> paragraphParser -> sectionParser -> headerFooterParser.
 */

import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import type { HeaderFooterType, HeaderReference, FooterReference } from "../types/document";
import type { ParseContext } from "./parseContext";
import { findChildren, getAttribute } from "./xmlParser";
import type { XmlElement } from "./xmlParser";

/**
 * Read a `w:type` attribute as one of ECMA-376's three `ST_HdrFtr` values.
 *
 * The enumeration is `even`, `default` and `first` (17.18.36); `default` is
 * the odd-page header, which is why a producer writing `odd` means the same
 * thing. Word opens such a package, so anything outside the enumeration reads
 * as the default the schema itself defaults to. Everything that compares
 * header/footer reference types has to read them through here, or a
 * normalisation at this boundary looks like a lost reference downstream.
 */
export function parseHeaderFooterType(
  typeAttr: string | null,
  context?: ParseContext,
): HeaderFooterType {
  switch (typeAttr) {
    case "first":
      return "first";
    case "even":
      return "even";
    case "default":
    case null:
      return "default";
    default:
      // `odd` is the common one: `default` is the odd-page header, so the
      // producer meant it, and Word opens the file. Reporting it keeps the
      // repack fidelity guard's comparison honest and tells a host that the
      // bytes it gets back will not match the bytes it gave.
      context?.warn({
        code: PARSE_WARNING_CODES.headerFooterTypeOutsideEnum,
        value: typeAttr,
        element: "w:type",
      });
      return "default";
  }
}

function parseHeaderFooterReference(element: XmlElement, context?: ParseContext) {
  const typeAttr = getAttribute(element, "w", "type");
  const rId = getAttribute(element, "r", "id") ?? "";

  return {
    type: parseHeaderFooterType(typeAttr, context?.scoped({ at: `r:id "${rId}"` })),
    rId,
  };
}

/**
 * Parse a header reference from sectPr (w:headerReference)
 */
export function parseHeaderReference(element: XmlElement, context?: ParseContext): HeaderReference {
  return parseHeaderFooterReference(element, context);
}

/**
 * Parse a footer reference from sectPr (w:footerReference)
 */
export function parseFooterReference(element: XmlElement, context?: ParseContext): FooterReference {
  return parseHeaderFooterReference(element, context);
}

/**
 * Parse all header references from a sectPr element
 */
export function parseHeaderReferences(sectPr: XmlElement): HeaderReference[] {
  const refs: HeaderReference[] = [];
  const headerRefElements = findChildren(sectPr, "w", "headerReference");

  for (const el of headerRefElements) {
    refs.push(parseHeaderReference(el));
  }

  return refs;
}

/**
 * Parse all footer references from a sectPr element
 */
export function parseFooterReferences(sectPr: XmlElement): FooterReference[] {
  const refs: FooterReference[] = [];
  const footerRefElements = findChildren(sectPr, "w", "footerReference");

  for (const el of footerRefElements) {
    refs.push(parseFooterReference(el));
  }

  return refs;
}
