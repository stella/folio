/**
 * Namespace resolution for parsed XML elements, shared by both XML parsers.
 *
 * Kept apart from `xmlParser.ts` so the streaming parser can attach the same
 * context without importing the module that dispatches to it.
 */

import type { XmlElement, XmlNamespaceScope } from "./xmlElement";

export const EMPTY_NAMESPACE_SCOPE: XmlNamespaceScope = { bindings: new Map() };

/**
 * The OOXML namespaces a WordprocessingML package binds, by their customary
 * prefix. The same table as `OOXML_NS` in `@stll/docx-utils`, so code reading
 * either agrees on every URI.
 */
export const OOXML_NS: {
  readonly w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  readonly a: "http://schemas.openxmlformats.org/drawingml/2006/main";
  readonly r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  readonly wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
  readonly wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
  readonly wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
  readonly wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
  readonly wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
  readonly pic: "http://schemas.openxmlformats.org/drawingml/2006/picture";
  readonly m: "http://schemas.openxmlformats.org/officeDocument/2006/math";
  readonly mc: "http://schemas.openxmlformats.org/markup-compatibility/2006";
  readonly v: "urn:schemas-microsoft-com:vml";
  readonly o: "urn:schemas-microsoft-com:office:office";
  readonly w14: "http://schemas.microsoft.com/office/word/2010/wordml";
  readonly w15: "http://schemas.microsoft.com/office/word/2012/wordml";
  readonly ct: "http://schemas.openxmlformats.org/package/2006/content-types";
  readonly pr: "http://schemas.openxmlformats.org/package/2006/relationships";
} = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  wpc: "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
  wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  w14: "http://schemas.microsoft.com/office/word/2010/wordml",
  w15: "http://schemas.microsoft.com/office/word/2012/wordml",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
  pr: "http://schemas.openxmlformats.org/package/2006/relationships",
};

export type OoxmlPrefix = keyof typeof OOXML_NS;

/**
 * Reserved by the XML Namespaces recommendation (§3): every document binds
 * this prefix to this URI whether or not it declares `xmlns:xml`, and no
 * document may bind it to anything else. A source that omits the declaration
 * — the overwhelming majority — must still resolve `xml:space`, `xml:lang`
 * and `xml:id` to it, or a reader mistakes the attribute for an unprefixed
 * one and a save drops its prefix.
 */
const XML_RESERVED_PREFIX = "xml";
export const XML_NAMESPACE_URI = "http://www.w3.org/XML/1998/namespace";

export const resolveNamespaceUri = (
  scope: XmlNamespaceScope | undefined,
  prefix: string,
): string | undefined => {
  let current = scope;
  while (current) {
    const value = current.bindings.get(prefix);
    if (value !== undefined) {
      return value;
    }
    current = current.parent;
  }
  return prefix === XML_RESERVED_PREFIX ? XML_NAMESPACE_URI : undefined;
};

/** Attach the element's resolved namespace metadata from its in-scope declarations. */
export const attachXmlNamespaceContext = (
  element: XmlElement,
  inheritedNamespaceScope: XmlNamespaceScope = EMPTY_NAMESPACE_SCOPE,
): XmlNamespaceScope => {
  let localBindings: Map<string, string> | null = null;
  if (element.attributes) {
    for (const [attribute, value] of Object.entries(element.attributes)) {
      if (
        typeof value !== "string" ||
        (attribute !== "xmlns" && !attribute.startsWith("xmlns:"))
      ) {
        continue;
      }
      localBindings ??= new Map();
      const prefix =
        attribute === "xmlns" ? "" : attribute.slice("xmlns:".length);
      localBindings.set(prefix, value);
    }
  }
  const namespaceScope =
    localBindings === null
      ? inheritedNamespaceScope
      : { bindings: localBindings, parent: inheritedNamespaceScope };

  Object.defineProperty(element, "namespaceScope", {
    configurable: false,
    enumerable: false,
    value: namespaceScope,
    writable: false,
  });

  const name = element.name ?? "";
  const colonIndex = name.indexOf(":");
  const prefix = colonIndex === -1 ? "" : name.slice(0, colonIndex);
  const namespaceUri = resolveNamespaceUri(namespaceScope, prefix);
  if (namespaceUri !== undefined) {
    Object.defineProperty(element, "namespaceUri", {
      configurable: false,
      enumerable: false,
      value: namespaceUri,
      writable: false,
    });
  }
  return namespaceScope;
};
