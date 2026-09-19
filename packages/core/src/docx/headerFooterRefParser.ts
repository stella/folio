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

/**
 * A reference with no `r:id` names no part, so it is not a reference.
 *
 * Coercing the missing attribute to `""` used to put the empty string in the
 * model, where it became a part-map key on one side and an `r:id=""` the
 * schema rejects on the other. Null here keeps the reference out of the model
 * entirely, which is what the source said.
 */
function parseHeaderFooterReference(element: XmlElement, context?: ParseContext) {
  const rId = getAttribute(element, "r", "id");
  if (rId === null || rId.length === 0) {
    return null;
  }

  return {
    type: parseHeaderFooterType(
      getAttribute(element, "w", "type"),
      context?.scoped({ at: `r:id "${rId}"` }),
    ),
    rId,
  };
}

/**
 * Parse a header reference from sectPr (w:headerReference)
 */
export function parseHeaderReference(
  element: XmlElement,
  context?: ParseContext,
): HeaderReference | null {
  return parseHeaderFooterReference(element, context);
}

/**
 * Parse a footer reference from sectPr (w:footerReference)
 */
export function parseFooterReference(
  element: XmlElement,
  context?: ParseContext,
): FooterReference | null {
  return parseHeaderFooterReference(element, context);
}

/**
 * Parse all header references from a sectPr element
 */
export function parseHeaderReferences(sectPr: XmlElement): HeaderReference[] {
  const refs: HeaderReference[] = [];
  const headerRefElements = findChildren(sectPr, "w", "headerReference");

  for (const el of headerRefElements) {
    const ref = parseHeaderReference(el);
    if (ref) {
      refs.push(ref);
    }
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
    const ref = parseFooterReference(el);
    if (ref) {
      refs.push(ref);
    }
  }

  return refs;
}
