/**
 * Namespace resolution for parsed XML elements, shared by both XML parsers.
 *
 * Kept apart from `xmlParser.ts` so the streaming parser can attach the same
 * context without importing the module that dispatches to it.
 */

import type { XmlElement, XmlNamespaceScope } from "./xmlParser";

export const EMPTY_NAMESPACE_SCOPE: XmlNamespaceScope = { bindings: new Map() };

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
  return undefined;
};

/** Attach the element's resolved namespace metadata from its in-scope declarations. */
export const attachXmlNamespaceContext = (
  element: XmlElement,
  inheritedNamespaceScope: XmlNamespaceScope = EMPTY_NAMESPACE_SCOPE,
): XmlNamespaceScope => {
  let localBindings: Map<string, string> | null = null;
  if (element.attributes) {
    for (const [attribute, value] of Object.entries(element.attributes)) {
      if (typeof value !== "string" || (attribute !== "xmlns" && !attribute.startsWith("xmlns:"))) {
        continue;
      }
      localBindings ??= new Map();
      const prefix = attribute === "xmlns" ? "" : attribute.slice("xmlns:".length);
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
