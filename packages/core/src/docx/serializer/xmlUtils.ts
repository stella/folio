/**
 * Shared XML utility functions for serializers.
 */

import { getLocalName, parseXml } from "../xmlParser";

const XML_SPECIAL_CHARACTER_PATTERN = /[&<>"']/u;
const XML_SPECIAL_CHARACTER_GLOBAL_PATTERN = /[&<>"']/gu;

export function escapeXml(text: string): string {
  if (!XML_SPECIAL_CHARACTER_PATTERN.test(text)) {
    return text;
  }
  return text.replace(XML_SPECIAL_CHARACTER_GLOBAL_PATTERN, (character) => {
    if (character === "&") {
      return "&amp;";
    }
    if (character === "<") {
      return "&lt;";
    }
    if (character === ">") {
      return "&gt;";
    }
    if (character === '"') {
      return "&quot;";
    }
    return "&apos;";
  });
}

/**
 * Format a numeric value as an integer XML attribute.
 *
 * OOXML measure types (twips, EMU, half-points, eighths-of-point) are
 * integer-typed in the schema (xs:unsignedInt / xs:long / xs:int). Word
 * rejects floating-point values (e.g. `0.7 * 1440 === 1008.0000000000001`
 * from `inches * TWIPS_PER_INCH`), even though tolerant readers accept them.
 * Coerce to a finite integer at every serialization site.
 *
 * `NaN`/`Infinity`/`null`/`undefined` collapse to `'0'` rather than leaking
 * literal `"NaN"` or `"Infinity"` into the XML.
 */
export function intAttr(value: number | undefined | null): string {
  if (value == null || !Number.isFinite(value)) {
    return "0";
  }
  return String(Math.round(value));
}

/**
 * Verify that `xml` parses to exactly one root element named `expectedLocalName`
 * (namespace prefix ignored). Used to gate raw XML snapshots — e.g. an SDT's
 * `rawPropertiesXml`/`rawEndPropertiesXml` — before splicing them verbatim
 * into a serialized document. Those snapshots are normally produced by our
 * own parser, but they can also arrive from an untrusted surface (a
 * programmatically constructed node, a collaboration payload); a value that
 * is not a single well-formed `<w:sdtPr>`/`<w:sdtEndPr>` element could inject
 * sibling markup or close the enclosing `<w:sdt>` early. Callers should fall
 * back to a synthesized properties block when this returns `false`.
 */
export function isSingleWellFormedElement(xml: string, expectedLocalName: string): boolean {
  const trimmed = xml.trim();
  if (!trimmed) {
    return false;
  }

  let parsed: ReturnType<typeof parseXml>;
  try {
    parsed = parseXml(trimmed);
  } catch {
    return false;
  }

  const roots = (parsed.elements ?? []).filter((element) => element.type === "element");
  if (roots.length !== 1) {
    return false;
  }

  const root = roots[0];
  return root?.name !== undefined && getLocalName(root.name) === expectedLocalName;
}
