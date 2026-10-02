// PARSE-WARNING-EXEMPT: package-local numeric identities carry no authored
// content; integer definitions and references are remapped together before
// parsing. Malformed lexical identities retain the parser's drop-and-warn path.

import {
  isValidOoxmlNumericId,
  ooxmlNumericIdDomain,
  assertValidOoxmlNumericId,
  mayContainOoxmlNumericIds,
  mayContainInvalidOoxmlNumericIds,
  isOoxmlNumericIdAttributeName,
} from "@stll/docx-core";
import { panic } from "better-result";

import { REVISION_ELEMENT_NAMES } from "./revisionIdNormalization";
import {
  scanStreamingXmlNumericIdAttributes,
  parseStreamingXmlWithIdentityVisitor,
} from "./streamingXmlParser";
import {
  getLocalName,
  getNamespaceUri,
  resolveAttributeNamespaceUri,
  type XmlElement,
} from "./xmlParser";
import { XmlResourceLimitError } from "./xmlResourceLimits";
import { replaceRawDocxXmlParts, type RawDocxContent } from "./unzip";

const identityKey = (value: string | number): string => {
  const spelling = String(value).trim();
  if (!/^[+-]?\d+$/u.test(spelling)) return spelling;
  const digits = spelling.replace(/^[+-]?0*/u, "");
  return digits === "" ? "0" : `${spelling.startsWith("-") ? "-" : ""}${digits}`;
};

/** Numeric schemas share a range, but note and numbering identities have separate namespaces. */
type IdentitySpaceOptions = {
  elementName: string;
  attributeName: string;
  domain: NonNullable<ReturnType<typeof ooxmlNumericIdDomain>>;
};

const identitySpace = ({ elementName, attributeName, domain }: IdentitySpaceOptions): string => {
  if (domain === "unsigned32") return "drawing";
  const element = getLocalName(elementName);
  const attribute = getLocalName(attributeName);
  switch (attribute === "val" ? element : attribute) {
    case "numId":
      return "numbering-instance";
    case "abstractNumId":
      return "numbering-abstract";
    case "numPicBulletId":
    case "lvlPicBulletId":
      return "numbering-picture";
    default:
      break;
  }
  switch (element) {
    case "footnote":
    case "footnoteReference":
      return "footnote";
    case "endnote":
    case "endnoteReference":
      return "endnote";
    case "id":
      return "content-control";
    default:
      return `annotation:${REVISION_ELEMENT_NAMES.has(element) ? "revision" : element.replace(/(?:Range)?(?:Start|End)$/u, "").replace(/^comment.*$/u, "comment")}`;
  }
};

const identifiedAttributes = (element: XmlElement) => {
  const attributes = [];
  for (const [name, value] of Object.entries(element.attributes ?? {})) {
    if (!isOoxmlNumericIdAttributeName({ elementName: element.name ?? "", attributeName: name }))
      continue;
    if (value === undefined || !/^[+-]?\d+$/u.test(String(value).trim())) continue;
    const domain = ooxmlNumericIdDomain({
      elementName: element.name ?? "",
      elementNamespace: getNamespaceUri(element),
      attributeName: name,
      attributeNamespace: resolveAttributeNamespaceUri(element, name),
    });
    if (domain === undefined) continue;
    attributes.push({
      name,
      value,
      domain,
      space: identitySpace({ elementName: element.name ?? "", attributeName: name, domain }),
    });
  }
  return attributes;
};

type IdentitySpace = { reserved: Set<number>; replacements: Map<string, string>; next: number };
type ImportedIdentitySpan = {
  start: number;
  end: number;
  space: string;
  key: string;
  element: XmlElement;
  attributeName: string;
};

type NumericIdNormalizationOptions = { onParsedDocument?: (document: XmlElement) => void };

/**
 * Repair out-of-range imported integers before any model or opaque XML is captured.
 * Reserve the complete package first, then remap each old value once per space.
 * Source splices preserve unrelated XML bytes; a valid package takes the fast path.
 */
export const normalizeImportedNumericIds = (
  parts: ReadonlyMap<string, string>,
  options: NumericIdNormalizationOptions = {},
): ReadonlyMap<string, string> => {
  if (![...parts.values()].some((xml) => mayContainInvalidOoxmlNumericIds(xml, "range")))
    return parts;
  const normalized = new Map(parts);
  const candidates = [...parts]
    .filter(([, xml]) => mayContainOoxmlNumericIds(xml))
    .toSorted(([left], [right]) => left.localeCompare(right));
  const spaces = new Map<string, IdentitySpace>();
  const reservedPools = new Map<string, Set<number>>();
  const changedSpans = new Map<string, ImportedIdentitySpan[]>();
  let parsedDocument: XmlElement | undefined;
  type ScanOptions = {
    path: string;
    xml: string;
    visitor: Parameters<typeof scanStreamingXmlNumericIdAttributes>[1];
  };
  const scan = ({ path, xml, visitor }: ScanOptions): void => {
    const scanned =
      path.toLowerCase() === "word/document.xml" && options.onParsedDocument !== undefined
        ? parseStreamingXmlWithIdentityVisitor(xml, visitor)
        : scanStreamingXmlNumericIdAttributes(xml, visitor);
    if (scanned.status === "unsupported") {
      throw new XmlResourceLimitError({
        message: `Numeric-id normalization could not safely scan ${path}`,
        limit: "syntax",
        observed: 0,
        allowed: 0,
      });
    }
    if (scanned.status === "parsed") parsedDocument = scanned.value;
  };
  for (const [path, xml] of candidates) {
    scan({
      path,
      xml,
      visitor: (element, attributeValueSpans) => {
        for (const attribute of identifiedAttributes(element)) {
          let space = spaces.get(attribute.space);
          if (space === undefined) {
            const poolName = attribute.space.startsWith("annotation:")
              ? "annotation"
              : attribute.space;
            let reserved = reservedPools.get(poolName);
            if (reserved === undefined) {
              reserved = new Set();
              reservedPools.set(poolName, reserved);
            }
            space = { reserved, replacements: new Map(), next: 1 };
            spaces.set(attribute.space, space);
          }
          if (isValidOoxmlNumericId(attribute.value, attribute.domain)) {
            space.reserved.add(Number(attribute.value));
            continue;
          }
          const key = identityKey(attribute.value);
          space.replacements.set(key, "");
          const span = attributeValueSpans.get(attribute.name);
          if (span === undefined) panic("Missing imported numeric identity source span");
          let spans = changedSpans.get(path);
          if (spans === undefined) {
            spans = [];
            changedSpans.set(path, spans);
          }
          spans.push({
            start: span.start,
            end: span.end,
            space: attribute.space,
            key,
            element,
            attributeName: attribute.name,
          });
        }
        return null;
      },
    });
  }
  for (const space of spaces.values()) {
    for (const key of space.replacements.keys()) {
      while (space.reserved.has(space.next)) space.next += 1;
      assertValidOoxmlNumericId({
        value: space.next,
        partPath: "import",
        elementName: "allocated identity",
        attributeName: "id",
      });
      const replacement = String(space.next);
      space.replacements.set(key, replacement);
      space.reserved.add(space.next);
      space.next += 1;
    }
  }
  for (const [path, spans] of changedSpans) {
    const xml = parts.get(path);
    if (xml === undefined) panic(`Missing imported numeric identity part: ${path}`);
    const chunks: string[] = [];
    let cursor = 0;
    // The reservation walk already resolved namespaces and decoded integer aliases.
    // Reuse its source spans rather than parsing changed parts a second time.
    for (const span of spans.toSorted((left, right) => left.start - right.start)) {
      const replacement = spaces.get(span.space)?.replacements.get(span.key);
      if (replacement === undefined || replacement === "")
        panic("Missing imported numeric identity replacement");
      if (span.element.attributes === undefined)
        panic("Missing imported numeric identity attributes");
      span.element.attributes[span.attributeName] = replacement;
      chunks.push(xml.slice(cursor, span.start), replacement);
      cursor = span.end;
    }
    chunks.push(xml.slice(cursor));
    normalized.set(path, chunks.join(""));
  }
  if (parsedDocument !== undefined) options.onParsedDocument?.(parsedDocument);
  return normalized;
};

/** Parser boundary: all captures and selective-save bytes see the same identities. */
export const normalizeRawDocxNumericIds = async (
  raw: RawDocxContent,
): Promise<XmlElement | undefined> => {
  let documentTree: XmlElement | undefined;
  const normalized = normalizeImportedNumericIds(raw.allXml, {
    onParsedDocument: (tree) => {
      documentTree = tree;
    },
  });
  if (normalized === raw.allXml) return documentTree;
  await replaceRawDocxXmlParts(raw, normalized);
  return documentTree;
};
