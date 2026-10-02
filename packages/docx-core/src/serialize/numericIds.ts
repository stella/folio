import { TaggedError } from "better-result";
import { XMLParser } from "fast-xml-parser";

import { MAX_REVISION_ID } from "../model/content";

const WORD_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);
const DRAWING_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://schemas.openxmlformats.org/drawingml/2006/picture",
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://purl.oclc.org/ooxml/drawingml/main",
  "http://purl.oclc.org/ooxml/drawingml/picture",
  "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
]);
const NUMBERING_ATTRIBUTES = new Set(["numId", "abstractNumId", "numPicBulletId"]);
const ID_VALUE_ELEMENTS = new Set(["id", "numId", "abstractNumId", "lvlPicBulletId"]);
const STRING_ID_ELEMENTS = new Set(["permStart", "permEnd"]);
const MAX_UNSIGNED_ID = 0xffff_ffff;
const MIN_SIGNED_ID = -0x8000_0000;

/** A writer would emit an identifier outside its OOXML schema's numeric domain. */
export class InvalidOoxmlNumericIdError extends TaggedError("InvalidOoxmlNumericIdError")<{
  message: string;
  partPath: string;
  elementName: string;
  attributeName: string;
  value: string;
}> {}

type NumericIdOptions = {
  value: string | number;
  partPath: string;
  elementName: string;
  attributeName: string;
  domain?: "signed32" | "unsigned32";
};

export const isValidOoxmlNumericId = (
  value: string | number,
  domain: "signed32" | "unsigned32" = "signed32",
): boolean => {
  const spelling = String(value).trim();
  const numeric = Number(spelling);
  const min = domain === "unsigned32" ? 0 : MIN_SIGNED_ID;
  const max = domain === "unsigned32" ? MAX_UNSIGNED_ID : MAX_REVISION_ID;
  return (
    /^[+-]?\d+$/u.test(spelling) && Number.isInteger(numeric) && numeric >= min && numeric <= max
  );
};

export const assertValidOoxmlNumericId = ({
  value,
  partPath,
  elementName,
  attributeName,
  domain = "signed32",
}: NumericIdOptions): void => {
  if (isValidOoxmlNumericId(value, domain)) return;
  const min = domain === "unsigned32" ? 0 : MIN_SIGNED_ID;
  const max = domain === "unsigned32" ? MAX_UNSIGNED_ID : MAX_REVISION_ID;
  throw new InvalidOoxmlNumericIdError({
    message: `${partPath}: ${elementName} ${attributeName} must be an integer in [${min}, ${max}], got ${String(value)}`,
    partPath,
    elementName,
    attributeName,
    value: String(value),
  });
};

type NumericIdAttributeOptions = {
  elementName: string;
  elementNamespace: string | undefined;
  attributeName: string;
  attributeNamespace: string | undefined;
};

/** Shared schema classification for import normalization and writer validation. */
export const ooxmlNumericIdDomain = ({
  elementName,
  elementNamespace,
  attributeName,
  attributeNamespace,
}: NumericIdAttributeOptions): "signed32" | "unsigned32" | undefined => {
  const elementLocalName = localName(elementName);
  const attributeLocalName = localName(attributeName);
  if (
    WORD_NAMESPACES.has(attributeNamespace ?? "") &&
    ((attributeLocalName === "id" &&
      !(WORD_NAMESPACES.has(elementNamespace ?? "") && STRING_ID_ELEMENTS.has(elementLocalName))) ||
      NUMBERING_ATTRIBUTES.has(attributeLocalName) ||
      (attributeLocalName === "val" &&
        WORD_NAMESPACES.has(elementNamespace ?? "") &&
        ID_VALUE_ELEMENTS.has(elementLocalName)))
  )
    return "signed32";
  if (
    attributeName === "id" &&
    DRAWING_NAMESPACES.has(elementNamespace ?? "") &&
    (elementLocalName === "docPr" || elementLocalName === "cNvPr")
  )
    return "unsigned32";
  return undefined;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const localName = (name: string): string => name.slice(name.indexOf(":") + 1);
const namespaceUri = (name: string, bindings: ReadonlyMap<string, string>): string | undefined => {
  const colon = name.indexOf(":");
  return bindings.get(colon === -1 ? "" : name.slice(0, colon));
};

// Decode raw attribute references once: an escaped ampersand must not turn
// literal entity text into a second character reference.
const decodeAttributeReferences = (value: string): string =>
  value.replace(
    /&(?:#(x[\da-f]+|\d+)|amp|lt|gt|quot|apos);/giu,
    (reference, digits: string | undefined) => {
      if (digits !== undefined) {
        const codePoint = digits.toLowerCase().startsWith("x")
          ? Number.parseInt(digits.slice(1), 16)
          : Number(digits);
        if (codePoint <= 0 || codePoint > 0x10ffff || (codePoint >= 0xd800 && codePoint <= 0xdfff))
          return reference;
        return String.fromCodePoint(codePoint);
      }
      switch (reference) {
        case "&amp;":
          return "&";
        case "&lt;":
          return "<";
        case "&gt;":
          return ">";
        case "&quot;":
          return '"';
        case "&apos;":
          return "'";
        default:
          return reference;
      }
    },
  );

const ID_ATTRIBUTE_NAMES = ["id", ...NUMBERING_ATTRIBUTES].join("|");
const ID_ELEMENT_NAMES = [...ID_VALUE_ELEMENTS].join("|");
type OoxmlNumericIdAttributeOptions = { elementName: string; attributeName: string };

/** Attributes an identity scan must retain, derived from the domain classifier. */
export const isOoxmlNumericIdAttributeName = ({
  elementName,
  attributeName,
}: OoxmlNumericIdAttributeOptions): boolean => {
  const attribute = localName(attributeName);
  return (
    attribute === "id" ||
    NUMBERING_ATTRIBUTES.has(attribute) ||
    (attribute === "val" && ID_VALUE_ELEMENTS.has(localName(elementName)))
  );
};
const ID_CANDIDATE = new RegExp(
  `\\b(?:${ID_ATTRIBUTE_NAMES})\\s*=|<(?:[^\\s<>/:]+:)?(?:${ID_ELEMENT_NAMES})(?:[\\s/>])`,
  "u",
);
/** Conservative presence check; safe IDs must still be reserved during import repair. */
export const mayContainOoxmlNumericIds = (xml: string): boolean => ID_CANDIDATE.test(xml);
const PREFIX = "[^\\s<>/:=\"']+";
const ID_ATTRIBUTE = new RegExp(
  `(?:^|\\s)((?:${PREFIX}:)?(?:${ID_ATTRIBUTE_NAMES}))\\s*=\\s*(["'])([\\s\\S]*?)\\2`,
  "gu",
);
const ID_VALUE_TAG = new RegExp(
  `<((?:${PREFIX}:)?(?:${ID_ELEMENT_NAMES}))(?=[\\s/>])((?:[^<>"']|"[^"]*"|'[^']*')*)>`,
  "gu",
);
const VALUE_ATTRIBUTE = new RegExp(
  `(?:^|\\s)((?:${PREFIX}:)?val)\\s*=\\s*(["'])([\\s\\S]*?)\\2`,
  "gu",
);
// Skip short, nonnegative integers lexically. Suspicious val attributes are
// intentionally not restricted to owner tags here; namespace/owner checks below
// distinguish identities from formatting values only when a candidate exists.
const RANGE_ATTRIBUTE = new RegExp(
  `(?:^|\\s)((?:${PREFIX}:)?(?:${ID_ATTRIBUTE_NAMES}|val))\\s*=\\s*(["'])\\s*([+-]?\\d{10,}|-\\d+|[^"'<>]*&[^"'<>]*)\\s*\\2`,
  "gu",
);
const hasRangeCandidate = (xml: string): boolean => {
  for (const match of xml.matchAll(RANGE_ATTRIBUTE)) {
    const value = match[3] ?? "";
    if (value.includes("&")) return true;
    const domain = (match[1] ?? "").includes(":") ? "signed32" : "unsigned32";
    if (!isValidOoxmlNumericId(value.trim(), domain)) return true;
  }
  return false;
};
const NAMESPACE_BINDING = new RegExp(
  `\\bxmlns(?::(${PREFIX}))?\\s*=\\s*(["'])([\\s\\S]*?)\\2`,
  "gu",
);
const prefixOf = (name: string): string => {
  const colon = name.indexOf(":");
  return colon === -1 ? "" : name.slice(0, colon);
};

/**
 * Conservative lexical preflight; namespace classification remains authoritative.
 * A prefix is considered numeric if any scope binds it to WordprocessingML. This can send
 * a shadowed foreign attribute to the full parser, but cannot hide a numeric one.
 * Suspicious values with entity-encoded bindings or values require the authoritative scan.
 */
export const mayContainInvalidOoxmlNumericIds = (
  xml: string,
  mode: "schema" | "range" = "schema",
): boolean => {
  if (mode === "range") {
    if (!hasRangeCandidate(xml)) return false;
  } else if (!mayContainOoxmlNumericIds(xml)) return false;
  const wordPrefixes = new Set<string>();
  for (const match of xml.matchAll(NAMESPACE_BINDING)) {
    const namespace = match[3] ?? "";
    if (namespace.includes("&")) return true;
    if (WORD_NAMESPACES.has(namespace)) wordPrefixes.add(match[1] ?? "");
  }
  const needsScan = (value: string, domain?: "signed32"): boolean => {
    if (value.includes("&")) return true;
    if (mode === "range" && !/^[+-]?\d+$/u.test(value.trim())) return false;
    // Only DrawingML numeric identities use unqualified id attributes.
    return !isValidOoxmlNumericId(value, domain ?? "unsigned32");
  };
  for (const match of xml.matchAll(ID_ATTRIBUTE)) {
    const name = match[1] ?? "";
    if (name.includes(":")) {
      if (wordPrefixes.has(prefixOf(name)) && needsScan(match[3] ?? "", "signed32")) return true;
      continue;
    }
    if (name === "id" && needsScan(match[3] ?? "")) return true;
  }
  for (const match of xml.matchAll(ID_VALUE_TAG)) {
    if (!wordPrefixes.has(prefixOf(match[1] ?? ""))) continue;
    for (const attribute of (match[2] ?? "").matchAll(VALUE_ATTRIBUTE)) {
      const name = attribute[1] ?? "";
      if (
        name.includes(":") &&
        wordPrefixes.has(prefixOf(name)) &&
        needsScan(attribute[3] ?? "", "signed32")
      )
        return true;
    }
  }
  return false;
};

/**
 * Validate both authored and opaque replayed identifiers by namespace URI.
 * This runs at ZIP exits too, so unchanged parts cannot bypass the writer guard.
 * Permission-range ids are strings in OOXML and intentionally remain unrestricted.
 */
export const assertValidOoxmlNumericIds = (xml: string, partPath: string): void => {
  if (!mayContainInvalidOoxmlNumericIds(xml)) return;
  const parser = new XMLParser({
    preserveOrder: true,
    ignoreAttributes: false,
    attributeNamePrefix: "",
    parseTagValue: false,
    parseAttributeValue: false,
    processEntities: false,
    attributeValueProcessor: (_name, value) => decodeAttributeReferences(value),
    htmlEntities: false,
    ignoreDeclaration: true,
  });
  const visit = (nodes: unknown, inherited: ReadonlyMap<string, string>): void => {
    if (!Array.isArray(nodes)) return;
    for (const node of nodes) {
      if (!isRecord(node)) continue;
      const attributes = node[":@"];
      let localBindings: Map<string, string> | undefined;
      if (isRecord(attributes)) {
        for (const [name, value] of Object.entries(attributes)) {
          if (typeof value !== "string") continue;
          if (name === "xmlns" || name.startsWith("xmlns:")) {
            localBindings ??= new Map(inherited);
            localBindings.set(name === "xmlns" ? "" : name.slice(6), value);
          }
        }
      }
      const bindings = localBindings ?? inherited;
      for (const [elementName, children] of Object.entries(node)) {
        if (elementName === ":@" || elementName.startsWith("#") || elementName.startsWith("?"))
          continue;
        const elementNamespace = namespaceUri(elementName, bindings);
        if (isRecord(attributes)) {
          for (const [attributeName, value] of Object.entries(attributes)) {
            if (typeof value !== "string" && typeof value !== "number") continue;
            const attributeNamespace = attributeName.includes(":")
              ? namespaceUri(attributeName, bindings)
              : undefined;
            const domain = ooxmlNumericIdDomain({
              elementName,
              elementNamespace,
              attributeName,
              attributeNamespace,
            });
            if (domain !== undefined) {
              assertValidOoxmlNumericId({ value, partPath, elementName, attributeName, domain });
            }
          }
        }
        visit(children, bindings);
      }
    }
  };
  const parsed: unknown = parser.parse(xml);
  visit(parsed, new Map());
};
