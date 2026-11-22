/**
 * OPC writes every relationships part in one namespace, in Strict and
 * Transitional packages alike. A package `_rels/.rels` whose root is not a
 * `Relationships` element in that namespace (one in the office relationships
 * namespace, the one `r:` attributes use, say) is not package relationships:
 * both engines refuse it whole, before any reader takes a guess at it.
 */
import { Result, TaggedError, type TaggedErrorClass } from "better-result";

import {
  getLocalName,
  getNamespaceUri,
  parseXmlDocument,
  type XmlElement,
} from "./xmlParser";

export const PACKAGE_RELATIONSHIPS_NAMESPACE_URI =
  "http://schemas.openxmlformats.org/package/2006/relationships";

/** The package relationships part's name, lowercased: OPC names are case-insensitive. */
export const PACKAGE_RELATIONSHIPS_PART = "_rels/.rels";

const MalformedPackageRelationshipsErrorBase: TaggedErrorClass<"MalformedPackageRelationshipsError"> =
  TaggedError("MalformedPackageRelationshipsError");

/** A package relationships part whose root is not package relationships. */
export class MalformedPackageRelationshipsError extends MalformedPackageRelationshipsErrorBase<{
  message: string;
  /** The part, as the archive spells it. */
  part: string;
  /** The root element's namespace URI; `undefined` when it has none. */
  namespaceUri: string | undefined;
}> {}

/** The root of the package relationships part `xml`, or its refusal. */
export const packageRelationshipsRoot = (
  part: string,
  xml: string,
): Result<XmlElement, MalformedPackageRelationshipsError> => {
  const root = parseXmlDocument(xml);
  const namespaceUri = root === null ? undefined : getNamespaceUri(root);
  if (
    root !== null &&
    getLocalName(root.name ?? "") === "Relationships" &&
    namespaceUri === PACKAGE_RELATIONSHIPS_NAMESPACE_URI
  )
    return Result.ok(root);
  return Result.err(
    new MalformedPackageRelationshipsError({
      message: `${part} is not package relationships: its root must be a Relationships element in ${PACKAGE_RELATIONSHIPS_NAMESPACE_URI}.`,
      part,
      namespaceUri,
    }),
  );
};
