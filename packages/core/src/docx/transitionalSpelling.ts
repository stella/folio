/**
 * Reading a Strict-spelled value as the Transitional number it stands for.
 *
 * Two places need it: the verbatim-capture conversion, which rewrites a
 * fragment on its way into a rebuilt part, and the typed projection, which has
 * to reach the same number so the model and the markup do not disagree about a
 * width. Both go through {@link transitionalSlotEncoding} rather than deciding
 * for themselves which attributes carry a unit.
 */

import {
  type SlotEncoding,
  TRANSITIONAL_NAMESPACE_BY_STRICT_URI,
  TRANSITIONAL_SLOT_ENCODINGS,
} from "./strictValueEncodings.gen";

/**
 * The Transitional URI a Strict one stands for, or the URI unchanged.
 *
 * Also used when a rebuilt part falls back to a binding its source declared:
 * declaring a Strict URI on a Transitional root puts every element under that
 * prefix back into the wrong vocabulary.
 */
export const toTransitionalNamespaceUri = (uri: string): string =>
  TRANSITIONAL_NAMESPACE_BY_STRICT_URI.get(uri) ?? uri;

/** True for a namespace ECMA-376 Part 4 republished under `purl.oclc.org`. */
export const isStrictNamespaceUri = (uri: string): boolean =>
  TRANSITIONAL_NAMESPACE_BY_STRICT_URI.has(uri);

/** Namespaces ECMA-376 Part 4 republished; nothing else starts with this. */
export const STRICT_URI_PREFIX = "http://purl.oclc.org/ooxml/";

/**
 * How one slot's Transitional type spells its value as a number.
 *
 * `namespaceUri` is the element's, in either conformance class; the lookup
 * translates it. Omit `attributeLocalName` for element text content.
 */
export const transitionalSlotEncoding = (
  namespaceUri: string | undefined,
  elementLocalName: string,
  attributeLocalName?: string,
): SlotEncoding | undefined => {
  if (namespaceUri === undefined) {
    return undefined;
  }
  const slot = `${toTransitionalNamespaceUri(namespaceUri)} ${elementLocalName}`;
  return TRANSITIONAL_SLOT_ENCODINGS.get(
    attributeLocalName === undefined ? slot : `${slot} @${attributeLocalName}`,
  );
};
