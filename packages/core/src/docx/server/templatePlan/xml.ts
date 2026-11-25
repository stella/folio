/**
 * Raw OOXML helpers shared by the template plan modules: element tests by
 * namespace and local name, source-range splicing, and the prefixes a part
 * binds, so generated markup uses the spelling the part already declares.
 */
import { escapeXmlAttribute, escapeXmlText } from "@stll/docx-core";
import {
  getXmlSourceRange,
  parseStreamingXmlWithSourceRanges,
} from "@stll/docx-core/xml";

import { spliceXml, type XmlSplice } from "../../selectiveXmlPatch";
import {
  getAttributeByNamespaceUri,
  getLocalName,
  NAMESPACES,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "../../xmlParser";

export { escapeXmlAttribute, escapeXmlText, spliceXml };

export const MARKUP_COMPATIBILITY_NAMESPACE: string = NAMESPACES.mc;
export const MATH_NAMESPACE: string = NAMESPACES.m;

/** Whether `element` is the WordprocessingML element `localName`. */
export const isW = (element: XmlElement, localName: string): boolean =>
  element.type === "element" &&
  getLocalName(element.name) === localName &&
  WORDPROCESSINGML_NAMESPACE_URIS.has(element.namespaceUri ?? "");

/** The WordprocessingML local name of `element`, or null for anything else. */
export const wName = (element: XmlElement): string | null =>
  element.type === "element" &&
  WORDPROCESSINGML_NAMESPACE_URIS.has(element.namespaceUri ?? "")
    ? getLocalName(element.name)
    : null;

export const elementChildren = (
  element: XmlElement | null | undefined,
): XmlElement[] =>
  (element?.elements ?? []).filter((child) => child.type === "element");

export const wChild = (
  element: XmlElement | null | undefined,
  localName: string,
): XmlElement | null =>
  elementChildren(element).find((child) => isW(child, localName)) ?? null;

export const wChildren = (
  element: XmlElement | null | undefined,
  localName: string,
): XmlElement[] =>
  elementChildren(element).filter((child) => isW(child, localName));

/** A WordprocessingML attribute, whatever prefix the part spells it with. */
export const wAttr = (
  element: XmlElement | null | undefined,
  localName: string,
): string | null =>
  getAttributeByNamespaceUri(
    element,
    WORDPROCESSINGML_NAMESPACE_URIS,
    localName,
  );

/** `w:val` of the child `localName`, or null when the child is absent. */
export const wChildVal = (
  element: XmlElement | null | undefined,
  localName: string,
): string | null => wAttr(wChild(element, localName), "val");

export const integerOrNull = (value: string | null): number | null => {
  if (value === null || !/^-?\d+$/u.test(value.trim())) {
    return null;
  }
  const parsed = Number.parseInt(value, 10);
  return Number.isSafeInteger(parsed) ? parsed : null;
};

/** The largest of `values`, or `floor` when none is larger. */
export const maxOf = (values: Iterable<number>, floor: number): number => {
  let max = floor;
  for (const value of values) {
    if (value > max) {
      max = value;
    }
  }
  return max;
};

/** A relationship attribute (`r:id`, `r:embed`, ...) by local name. */
export const relationshipAttr = (
  element: XmlElement,
  localName: string,
): string | null =>
  getAttributeByNamespaceUri(
    element,
    OFFICE_RELATIONSHIP_NAMESPACE_URIS,
    localName,
  );

export type XmlRange = { start: number; end: number };
export type { XmlSplice };

/**
 * The slice `[start, end)` of `xml` with the splices that lie inside it
 * applied; null when the owner of part splices refuses the result.
 */
export const sliceWithSplices = (
  xml: string,
  { start, end }: XmlRange,
  splices: readonly XmlSplice[],
): string | null =>
  spliceXml(
    xml.slice(start, end),
    splices
      .filter((splice) => splice.start >= start && splice.end <= end)
      .map((splice) => ({
        start: splice.start - start,
        end: splice.end - start,
        newXml: splice.newXml,
      })),
  );

export type RawPart = {
  path: string;
  xml: string;
  /** The root element, parsed with source ranges. */
  root: XmlElement;
};

/** Parse a part keeping every element's source range, or null when it is not XML. */
export const readRawPart = (path: string, xml: string): RawPart | null => {
  const parsed = parseStreamingXmlWithSourceRanges(xml);
  if (parsed.status !== "parsed") {
    return null;
  }
  const root = elementChildren(parsed.value).at(0);
  return root === undefined ? null : { path, xml, root };
};

export const rangeOf = (element: XmlElement): XmlRange | null =>
  getXmlSourceRange(element) ?? null;

/**
 * Where the content of `element` sits in its source: after the opening tag
 * and before the closing one. Null for an element written self-closed.
 */
export const innerRangeOf = (
  xml: string,
  element: XmlElement,
): XmlRange | null => {
  const range = rangeOf(element);
  if (range === null) {
    return null;
  }
  if (xml.slice(range.end - 2, range.end) === "/>") {
    return null;
  }
  const close = xml.lastIndexOf("</", range.end - 1);
  let open = -1;
  let quote = "";
  for (let index = range.start; index < range.end; index += 1) {
    const char = xml[index];
    if (quote !== "") {
      if (char === quote) {
        quote = "";
      }
    } else if (char === '"' || char === "'") {
      quote = char;
    } else if (char === ">") {
      open = index + 1;
      break;
    }
  }
  return open === -1 || close < open ? null : { start: open, end: close };
};

/** The prefix `root` binds to one of `uris`, or null when it binds none. */
export const prefixFor = (
  root: XmlElement,
  uris: ReadonlySet<string>,
): { prefix: string; uri: string } | null => {
  for (const [name, value] of Object.entries(root.attributes ?? {})) {
    if (
      name.startsWith("xmlns:") &&
      value !== undefined &&
      uris.has(String(value))
    ) {
      return { prefix: name.slice("xmlns:".length), uri: String(value) };
    }
  }
  return null;
};

/** The prefixes generated markup uses inside one part. */
export type PartPrefixes = {
  /** WordprocessingML prefix, e.g. `w`. */
  w: string;
  /** WordprocessingML namespace URI of the part. */
  wUri: string;
  /** Relationship attribute prefix and the declaration it needs, when the root has none. */
  r: string;
  rDeclaration: string;
};

const FALLBACK_RELATIONSHIP_PREFIX = "folioR";

/** Null when the part's root binds no prefix to WordprocessingML. */
export const partPrefixes = (root: XmlElement): PartPrefixes | null => {
  const w = prefixFor(root, WORDPROCESSINGML_NAMESPACE_URIS);
  if (w === null) {
    return null;
  }
  const r = prefixFor(root, OFFICE_RELATIONSHIP_NAMESPACE_URIS);
  return {
    w: w.prefix,
    wUri: w.uri,
    r: r?.prefix ?? FALLBACK_RELATIONSHIP_PREFIX,
    rDeclaration:
      r === null
        ? ` xmlns:${FALLBACK_RELATIONSHIP_PREFIX}="${NAMESPACES.r}"`
        : "",
  };
};

/** Visit every element under `root` in document order, `root` included. */
export const walkElements = (
  root: XmlElement,
  visit: (element: XmlElement) => void,
): void => {
  const stack: XmlElement[] = [root];
  while (stack.length > 0) {
    const element = stack.pop();
    if (element === undefined || element.type !== "element") {
      continue;
    }
    visit(element);
    const children = element.elements ?? [];
    for (let index = children.length - 1; index >= 0; index -= 1) {
      const child = children[index];
      if (child !== undefined) {
        stack.push(child);
      }
    }
  }
};
