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

export const assertValidOoxmlNumericId = ({
  value,
  partPath,
  elementName,
  attributeName,
  domain = "signed32",
}: NumericIdOptions): void => {
  const spelling = String(value).trim();
  const numeric = Number(spelling);
  const min = domain === "unsigned32" ? 0 : MIN_SIGNED_ID;
  const max = domain === "unsigned32" ? MAX_UNSIGNED_ID : MAX_REVISION_ID;
  if (
    /^[+-]?\d+$/u.test(spelling) &&
    Number.isInteger(numeric) &&
    numeric >= min &&
    numeric <= max
  ) {
    return;
  }
  throw new InvalidOoxmlNumericIdError({
    message: `${partPath}: ${elementName} ${attributeName} must be an integer in [${min}, ${max}], got ${String(value)}`,
    partPath,
    elementName,
    attributeName,
    value: String(value),
  });
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const localName = (name: string): string => name.slice(name.indexOf(":") + 1);
const namespaceUri = (name: string, bindings: ReadonlyMap<string, string>): string | undefined => {
  const colon = name.indexOf(":");
  return bindings.get(colon === -1 ? "" : name.slice(0, colon));
};

/**
 * Validate both authored and opaque replayed identifiers by namespace URI.
 * This runs at ZIP exits too, so unchanged parts cannot bypass the writer guard.
 * Permission-range ids are strings in OOXML and intentionally remain unrestricted.
 */
export const assertValidOoxmlNumericIds = (xml: string, partPath: string): void => {
  if (
    !/\b(?:id|numId|abstractNumId|numPicBulletId)\s*=|<(?:[^\s<>/:]+:)?(?:id|numId|abstractNumId|lvlPicBulletId)(?:[\s/>])/u.test(
      xml,
    )
  )
    return;
  const parser = new XMLParser({
    preserveOrder: true,
    ignoreAttributes: false,
    attributeNamePrefix: "",
    parseTagValue: false,
    parseAttributeValue: false,
    processEntities: true,
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
        const elementLocalName = localName(elementName);
        const elementNamespace = namespaceUri(elementName, bindings);
        if (isRecord(attributes)) {
          for (const [attributeName, value] of Object.entries(attributes)) {
            if (typeof value !== "string" && typeof value !== "number") continue;
            const attributeLocalName = localName(attributeName);
            const attributeNamespace = attributeName.includes(":")
              ? namespaceUri(attributeName, bindings)
              : undefined;
            const wordId =
              WORD_NAMESPACES.has(attributeNamespace ?? "") &&
              ((attributeLocalName === "id" &&
                !(
                  WORD_NAMESPACES.has(elementNamespace ?? "") &&
                  STRING_ID_ELEMENTS.has(elementLocalName)
                )) ||
                NUMBERING_ATTRIBUTES.has(attributeLocalName) ||
                (attributeLocalName === "val" &&
                  WORD_NAMESPACES.has(elementNamespace ?? "") &&
                  ID_VALUE_ELEMENTS.has(elementLocalName)));
            const drawingId =
              attributeName === "id" &&
              DRAWING_NAMESPACES.has(elementNamespace ?? "") &&
              (elementLocalName === "docPr" || elementLocalName === "cNvPr");
            if (wordId || drawingId) {
              assertValidOoxmlNumericId({
                value,
                partPath,
                elementName,
                attributeName,
                domain: drawingId ? "unsigned32" : "signed32",
              });
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
