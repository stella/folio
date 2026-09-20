/**
 * The one reader for "the paraId written on this element".
 *
 * A `paraId` is a join key: `w15:commentEx/@w15:paraId`,
 * `w16cex:comment/@w16cex:paraId` and `w15:paraIdParent` all name a paragraph
 * through the `w14:paraId` it carries. Three Word generations write the same
 * key under three prefixes, and a producer is free to bind any of those
 * namespaces to a prefix of its own choosing, so the key must be resolved by
 * namespace URI. Resolving it by local name alone is worse than missing it: an
 * unrelated `vendor:paraId` would key a comment into another thread and carry
 * that thread's date, parent and resolved state.
 *
 * Every reader of the key calls this, so no call site can come to disagree
 * with another about which attribute is a paraId.
 */

import { OOXML_NAMESPACES } from "./serializer/partNamespaces";
import {
  NAMESPACES,
  WORDPROCESSINGML_NAMESPACE_URIS,
  getAttributeByNamespaceUri,
  type XmlElement,
} from "./xmlParser";

/**
 * The namespaces a paraId-valued attribute may be written in: the Word 2010,
 * 2012 and 2018 wordml extensions.
 *
 * The 2018 URI is read from the serializer's table rather than retyped, so the
 * spelling folio writes and the spelling it accepts cannot drift.
 */
export const PARA_ID_NAMESPACE_URIS: ReadonlySet<string> = new Set([
  NAMESPACES.w14,
  NAMESPACES.w15,
  OOXML_NAMESPACES.w16cex.uri,
]);

/**
 * The conventional prefixes, read literally as a last resort.
 *
 * A part that writes `w14:paraId` without declaring `xmlns:w14` is malformed,
 * and no URI lookup can resolve it; Word reads those files, so folio does too.
 */
const CONVENTIONAL_PARA_ID_PREFIXES = ["w14", "w15", "w16cex"] as const;

const paraIdValuedAttribute = (element: XmlElement, localName: string): string | undefined => {
  const byUri = getAttributeByNamespaceUri(element, PARA_ID_NAMESPACE_URIS, localName);
  if (byUri !== null) {
    return byUri;
  }

  for (const prefix of CONVENTIONAL_PARA_ID_PREFIXES) {
    const literal = element.attributes?.[`${prefix}:${localName}`];
    if (literal !== undefined) {
      return String(literal);
    }
  }

  const inWordprocessingml = getAttributeByNamespaceUri(
    element,
    WORDPROCESSINGML_NAMESPACE_URIS,
    localName,
  );
  if (inWordprocessingml !== null) {
    return inWordprocessingml;
  }

  const literalW = element.attributes?.[`w:${localName}`];
  return literalW === undefined ? undefined : String(literalW);
};

/** The paraId this element carries, or `undefined` when it carries none. */
export const paraIdAttribute = (element: XmlElement): string | undefined =>
  paraIdValuedAttribute(element, "paraId");

/** The paraId of the parent this element names, by the same rule. */
export const paraIdParentAttribute = (element: XmlElement): string | undefined =>
  paraIdValuedAttribute(element, "paraIdParent");

/**
 * The textId this element carries: a paragraph identity written beside
 * `paraId`, in the same namespaces and by the same rule.
 */
export const textIdAttribute = (element: XmlElement): string | undefined =>
  paraIdValuedAttribute(element, "textId");
