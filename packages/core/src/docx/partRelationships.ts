/**
 * The relationship a part needs to another part, found by its type and the
 * part it reaches, and added or removed without rewriting the authored
 * entries around it.
 */
import { panic } from "better-result";

import { escapeXmlAttribute } from "@stll/docx-core";
import {
  getXmlSourceRange,
  parseStreamingXmlWithSourceRanges,
} from "@stll/docx-core/xml";

import { resolvePackageRelationshipTarget } from "./packageParts";
import { PACKAGE_RELATIONSHIPS_NAMESPACE_URI } from "./packageRelationshipsRoot";
import { parseRelationships } from "./relsParser";
import { spliceXml } from "./selectiveXmlPatch";
import {
  getAttribute,
  getChildElements,
  getLocalName,
  getNamespacePrefix,
  getNamespaceUri,
  parseXmlDocument,
  type XmlElement,
} from "./xmlParser";
import { appendToXmlRoot } from "./xmlRootInsertion";

type PartRelationship = {
  /** The relationships part, `word/_rels/document.xml.rels` say. */
  relsPath: string;
  type: string;
  /** The part the relationship reaches, as the archive names it. */
  partPath: string;
};

const reaches = (
  { relsPath, partPath }: PartRelationship,
  target: string,
): boolean =>
  resolvePackageRelationshipTarget(target, relsPath) === partPath.toLowerCase();

const isRelationshipsRoot = (element: XmlElement): boolean =>
  getLocalName(element.name) === "Relationships" &&
  getNamespaceUri(element) === PACKAGE_RELATIONSHIPS_NAMESPACE_URI;

/**
 * `xml` with the relationship; as it is when it holds it already, and `null`
 * when it is no relationships part.
 */
export const withPartRelationship = (
  xml: string,
  relationship: PartRelationship,
): string | null => {
  const root = parseXmlDocument(xml);
  if (root?.name === undefined || !isRelationshipsRoot(root)) return null;
  const relationships = parseRelationships(xml);
  let next = 1;
  for (const { id, type, target, targetMode } of relationships.values()) {
    if (
      type === relationship.type &&
      targetMode !== "External" &&
      reaches(relationship, target)
    )
      return xml;
    const number = /^rId(?<number>\d+)$/u.exec(id)?.groups?.["number"];
    if (number !== undefined) next = Math.max(next, Number(number) + 1);
  }
  const prefix = getNamespacePrefix(root.name);
  const name = prefix ? `${prefix}:Relationship` : "Relationship";
  const owner = relationship.relsPath.replace(/_rels\/[^/]*$/u, "");
  const target = relationship.partPath.startsWith(owner)
    ? relationship.partPath.slice(owner.length)
    : `/${relationship.partPath}`;
  return appendToXmlRoot(xml, {
    rootName: root.name,
    markup: `<${name} Id="rId${next}" Type="${escapeXmlAttribute(relationship.type)}" Target="${escapeXmlAttribute(target)}"/>`,
  });
};

/** `xml` without the relationships of the type that reach the part. */
export const withoutPartRelationship = (
  xml: string,
  relationship: PartRelationship,
): string => {
  const parsed = parseStreamingXmlWithSourceRanges(xml);
  const root =
    parsed.status === "parsed"
      ? getChildElements(parsed.value).find(isRelationshipsRoot)
      : undefined;
  if (root === undefined) return xml;
  const relationships = parseRelationships(xml);
  const splices = getChildElements(root).flatMap((element) => {
    const id = getAttribute(element, null, "Id");
    const target = getAttribute(element, null, "Target");
    if (
      id === null ||
      target === null ||
      relationships.get(id)?.type !== relationship.type ||
      !reaches(relationship, target)
    )
      return [];
    const range =
      getXmlSourceRange(element) ??
      panic("A parsed relationship must have a source range.");
    return [{ ...range, newXml: "" }];
  });
  return (
    spliceXml(xml, splices) ??
    panic("A relationship removal must preserve XML ranges.")
  );
};
