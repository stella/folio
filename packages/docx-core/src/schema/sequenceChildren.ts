/**
 * The one writer of a property set's children, in the order the schema
 * declares them.
 *
 * `w:rPr`, `w:tblPr` and `w:sectPr` each declare their children in one order,
 * and folio writes them in it. How binding that order is differs by container:
 * `CT_TblPrBase` is an `xsd:sequence` of distinct names and a validating
 * consumer refuses a `w:tblPr` written in any other order, while `EG_RPrBase`
 * is an `xsd:choice` referenced `maxOccurs="unbounded"`, so a `w:rPr` in any
 * order is valid. Writing the canonical order anyway is what keeps two
 * serializers of the same element from disagreeing, and it is what Word
 * writes.
 *
 * The order is read from the generated list rather than restated as the order
 * of a list of `if` statements — the restatement is what drifted, twice: once
 * inside folio-core, and once as a second `w:rPr` writer in this package that
 * put `w:sz` and `w:highlight` before `w:rFonts`.
 */

import type { PreservedMarkup } from "../model/preservedMarkup";
import { OOXML_NS } from "@stll/docx-utils";
import { XMLParser, XMLValidator } from "fast-xml-parser";
import { TaggedError } from "better-result";
import {
  SEQUENCE_CHILDREN,
  type SequenceChild,
  type SequenceContainer,
} from "./sequenceChildren.gen";

export {
  SEQUENCE_CHILDREN,
  type SequenceChild,
  type SequenceContainer,
} from "./sequenceChildren.gen";

type SerializeSequenceChildrenOptions<Container extends SequenceContainer> = {
  container: Container;
  /**
   * The modelled children keyed by element name rather than pre-ordered, so
   * the caller cannot state an order of its own.
   *
   * An empty string means the caller wrote nothing for that child.
   */
  modelled: ReadonlyArray<readonly [name: SequenceChild<Container>, xml: string]>;
  /**
   * The container's unmodelled markup, each capture carrying the schema
   * ordinal it was read at, so it lands between the same two neighbours.
   *
   * A caller that builds a document from scratch has none.
   */
  preserved?: PreservedMarkup | undefined;
};

const MAX_PRESERVED_CHILD_XML_CHARACTERS = 1024 * 1024;
const MAX_PRESERVED_CHILD_COUNT = 4096;
const MAX_PRESERVED_MARKUP_CHARACTERS = 4 * 1024 * 1024;
const PRESERVED_CHILD_WRAPPER = "folio-preserved-child-root";
const preservedChildParser = new XMLParser({
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
  ignoreDeclaration: false,
  ignorePiTags: false,
  processEntities: false,
  htmlEntities: false,
});

class InvalidPreservedChildXmlError extends TaggedError("InvalidPreservedChildXmlError")<{
  message: string;
  index: number;
}> {}

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const XML_NAMESPACE_URI = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE_URI = "http://www.w3.org/2000/xmlns/";
const PREDEFINED_XML_ENTITIES = new Set(["amp", "apos", "gt", "lt", "quot"]);
const INHERITED_NAMESPACE_BINDINGS: ReadonlyMap<string, string> = new Map([
  ...Object.entries(OOXML_NS),
  ["xml", XML_NAMESPACE_URI],
]);

const isXmlCharacter = (codePoint: number): boolean =>
  codePoint === 0x9 ||
  codePoint === 0xa ||
  codePoint === 0xd ||
  (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
  (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
  (codePoint >= 0x10000 && codePoint <= 0x10ffff);

const isValidXmlReference = (reference: string): boolean => {
  if (PREDEFINED_XML_ENTITIES.has(reference)) {
    return true;
  }
  if (/^#[0-9]+$/u.test(reference)) {
    return isXmlCharacter(Number(reference.slice(1)));
  }
  if (/^#x[0-9a-fA-F]+$/u.test(reference)) {
    return isXmlCharacter(Number.parseInt(reference.slice(2), 16));
  }
  return false;
};

/** Validate references only where XML treats them as references. */
const hasOnlyValidXmlReferences = (xml: string): boolean => {
  let index = 0;
  while (index < xml.length) {
    if (xml.startsWith("<![CDATA[", index)) {
      const end = xml.indexOf("]]>", index + 9);
      if (end === -1) {
        return false;
      }
      index = end + 3;
      continue;
    }
    if (xml.startsWith("<!--", index)) {
      const end = xml.indexOf("-->", index + 4);
      if (end === -1) {
        return false;
      }
      index = end + 3;
      continue;
    }
    if (xml.startsWith("<?", index)) {
      const end = xml.indexOf("?>", index + 2);
      if (end === -1) {
        return false;
      }
      index = end + 2;
      continue;
    }
    if (xml[index] !== "&") {
      index += 1;
      continue;
    }
    const end = xml.indexOf(";", index + 1);
    if (end === -1 || !isValidXmlReference(xml.slice(index + 1, end))) {
      return false;
    }
    index = end + 1;
  }
  return true;
};

const namespacePrefix = (qualifiedName: string): string | undefined => {
  const colon = qualifiedName.indexOf(":");
  return colon > 0 ? qualifiedName.slice(0, colon) : undefined;
};

const applyNamespaceDeclarations = (
  attributes: Readonly<Record<string, unknown>>,
  inherited: ReadonlyMap<string, string>,
): ReadonlyMap<string, string> | undefined => {
  let bindings: Map<string, string> | undefined;
  for (const [name, rawValue] of Object.entries(attributes)) {
    if (name !== "xmlns" && !name.startsWith("xmlns:")) {
      continue;
    }
    if (typeof rawValue !== "string") {
      return undefined;
    }
    const prefix = name === "xmlns" ? "" : name.slice("xmlns:".length);
    const lowerPrefix = prefix.toLowerCase();
    if (
      prefix === "xmlns" ||
      (lowerPrefix.startsWith("xml") && prefix !== "xml") ||
      (prefix === "xml" && rawValue !== XML_NAMESPACE_URI) ||
      (prefix !== "xml" && rawValue === XML_NAMESPACE_URI) ||
      rawValue === XMLNS_NAMESPACE_URI ||
      (prefix !== "" && rawValue === "")
    ) {
      return undefined;
    }
    const protectedBinding = INHERITED_NAMESPACE_BINDINGS.get(prefix);
    if (protectedBinding !== undefined && protectedBinding !== rawValue) {
      return undefined;
    }
    bindings ??= new Map(inherited);
    bindings.set(prefix, rawValue);
  }
  return bindings ?? inherited;
};

const hasValidNamespaces = (
  root: Record<string, unknown>,
  inherited: ReadonlyMap<string, string>,
): boolean => {
  const pending: Array<{
    node: Record<string, unknown>;
    bindings: ReadonlyMap<string, string>;
  }> = [{ node: root, bindings: inherited }];
  while (pending.length > 0) {
    const current = pending.pop();
    if (current === undefined) {
      continue;
    }
    const { node, bindings: parentBindings } = current;
    const nodeNames = Object.keys(node).filter((name) => name !== ":@");
    const name = nodeNames[0];
    if (nodeNames.length !== 1 || name === undefined) {
      return false;
    }
    if (name.startsWith("#") || name.startsWith("?") || name.startsWith("!")) {
      continue;
    }

    const rawAttributes = node[":@"];
    const attributes = isRecord(rawAttributes) ? rawAttributes : {};
    const bindings = applyNamespaceDeclarations(attributes, parentBindings);
    if (bindings === undefined) {
      return false;
    }
    const elementPrefix = namespacePrefix(name);
    if (elementPrefix !== undefined && !bindings.has(elementPrefix)) {
      return false;
    }
    const expandedAttributeNames = new Set<string>();
    for (const attributeName of Object.keys(attributes)) {
      if (attributeName === "xmlns" || attributeName.startsWith("xmlns:")) {
        continue;
      }
      const attributePrefix = namespacePrefix(attributeName);
      if (attributePrefix !== undefined && !bindings.has(attributePrefix)) {
        return false;
      }
      const namespace = attributePrefix === undefined ? "" : bindings.get(attributePrefix);
      const colon = attributeName.indexOf(":");
      const localName = colon === -1 ? attributeName : attributeName.slice(colon + 1);
      const expandedName = `${namespace ?? ""}\u0000${localName}`;
      if (expandedAttributeNames.has(expandedName)) {
        return false;
      }
      expandedAttributeNames.add(expandedName);
    }

    const children = node[name];
    if (!Array.isArray(children)) {
      return false;
    }
    for (const child of children) {
      if (!isRecord(child)) {
        return false;
      }
      pending.push({ node: child, bindings });
    }
  }
  return true;
};

/** Whether a captured child is exactly one bounded, well-formed XML element. */
export const isSafePreservedChildXml = (xml: string): boolean => {
  const trimmed = xml.trim();
  if (
    trimmed.length === 0 ||
    trimmed.length > MAX_PRESERVED_CHILD_XML_CHARACTERS ||
    /<!\s*(?:DOCTYPE|ENTITY)\b/iu.test(trimmed) ||
    !hasOnlyValidXmlReferences(trimmed)
  ) {
    return false;
  }

  const wrapped = `<${PRESERVED_CHILD_WRAPPER}>${trimmed}</${PRESERVED_CHILD_WRAPPER}>`;
  if (XMLValidator.validate(wrapped) !== true) {
    return false;
  }

  let parsed: unknown;
  try {
    parsed = preservedChildParser.parse(wrapped);
  } catch {
    return false;
  }
  if (!Array.isArray(parsed) || parsed.length !== 1) {
    return false;
  }
  const wrapper = parsed[0];
  if (!isRecord(wrapper)) {
    return false;
  }
  const children = wrapper[PRESERVED_CHILD_WRAPPER];
  if (!Array.isArray(children) || children.length !== 1) {
    return false;
  }
  const child = children[0];
  if (!isRecord(child)) {
    return false;
  }
  const nodeNames = Object.keys(child).filter((name) => name !== ":@");
  const nodeName = nodeNames[0];
  return (
    nodeNames.length === 1 &&
    nodeName !== undefined &&
    !nodeName.startsWith("#") &&
    !nodeName.startsWith("?") &&
    !nodeName.startsWith("!") &&
    hasValidNamespaces(child, INHERITED_NAMESPACE_BINDINGS)
  );
};

/** Whether a capture collection stays within its aggregate resource budget. */
export const isWithinPreservedMarkupBudget = (children: readonly unknown[]): boolean => {
  if (children.length > MAX_PRESERVED_CHILD_COUNT) {
    return false;
  }
  let characters = 0;
  for (const child of children) {
    if (!isRecord(child)) {
      continue;
    }
    const xml = child["xml"];
    if (typeof xml !== "string") {
      continue;
    }
    characters += xml.length;
    if (characters > MAX_PRESERVED_MARKUP_CHARACTERS) {
      return false;
    }
  }
  return true;
};

/** Reject unsafe captured children before a caller performs any transformation. */
export const assertSafePreservedMarkup = (preserved: PreservedMarkup | undefined): void => {
  const children = preserved?.children ?? [];
  if (!isWithinPreservedMarkupBudget(children)) {
    throw new InvalidPreservedChildXmlError({
      message: "Preserved container markup exceeds its aggregate resource budget",
      index: -1,
    });
  }
  for (const { index, xml } of children) {
    if (!isSafePreservedChildXml(xml)) {
      throw new InvalidPreservedChildXmlError({
        message: "Preserved container markup must be one bounded, well-formed XML element",
        index,
      });
    }
  }
};

/**
 * A property set's children, modelled and captured, in schema order.
 *
 * Ties are only possible between a declared child and a capture sharing its
 * slot — an undeclared name takes the place of the last declared child before
 * it. The sort is stable and the modelled child leads.
 */
export const serializeSequenceChildren = <Container extends SequenceContainer>({
  container,
  modelled,
  preserved,
}: SerializeSequenceChildrenOptions<Container>): string[] => {
  const declared: readonly string[] = SEQUENCE_CHILDREN[container];
  const placed: Array<{ at: number; rank: number; xml: string }> = [];
  for (const [name, xml] of modelled) {
    if (xml.length > 0) {
      placed.push({ at: declared.indexOf(name), rank: 0, xml });
    }
  }
  const preservedChildren = preserved?.children ?? [];
  assertSafePreservedMarkup(preserved);
  for (const { index, xml } of preservedChildren) {
    placed.push({ at: index, rank: 1, xml });
  }
  return placed
    .sort((left, right) => left.at - right.at || left.rank - right.rank)
    .map(({ xml }) => xml);
};
