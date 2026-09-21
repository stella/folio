/**
 * The attribute remainder: what an element carried and its parser has no field for.
 *
 * The child dispatcher makes a container's unrecognised *children* a decision
 * somebody wrote down. Its attributes had no such branch at all: a parser reads
 * the ones it models off the element and the serializer rebuilds the start tag
 * from the model, so everything else goes. Word writes a revision-session id
 * (`w:rsidR` and its family) on nearly every paragraph, run, row and section,
 * and a save rewrote the whole document's revision history.
 *
 * Two rules make this safe, and both are the opposite of what a name-matching
 * remainder would do:
 *
 * - **Resolve, then decide.** folio reads an attribute by namespace URI plus
 *   local name, with a local-name fallback across prefixes, so an element that
 *   binds a second prefix to the WordprocessingML URI and writes `altw:rsidR`
 *   is read exactly as `w:rsidR` is. A remainder that matched the spelling
 *   would keep a second copy of an attribute the parser had already read. So
 *   the decision is made on the resolved local name, and what is written back
 *   is the canonical prefix for the resolved URI, never the source's.
 * - **The writer, not the predicate, prevents a double.** The modelled
 *   attributes are handed to {@link serializePreservedAttributes} as the
 *   fragments it is about to emit, and a remainder entry that would spell one
 *   of them again is dropped. The predicate below mirrors what each parser
 *   reads, and a mirror drifts; this is the guard that keeps the drift from
 *   reaching the part as a duplicate attribute, which makes it unopenable.
 *
 * A namespace declaration is not in the remainder. `partNamespaces.ts` derives
 * a rebuilt part's bindings from the prefixes the part actually uses, so
 * replaying the source element's own `xmlns:*` would fight the root's. For the
 * same reason an attribute whose namespace that table cannot spell is not kept:
 * writing it would emit an unbound prefix and fail the save, which is a worse
 * answer than recording the drop in the container contract.
 */

import type { PreservedAttribute } from "@stll/docx-core/model";

import { escapeXmlAttribute } from "@stll/docx-core";
import { TaggedError } from "better-result";

import { type ModelledAttributes, modelledAttributeNames } from "./propertyElementAttributes";
import { OOXML_NAMESPACES } from "./serializer/partNamespaces";
import { getLocalName, resolveAttributeNamespaceUri, type XmlElement } from "./xmlParser";

/** Namespace URI to the prefix a rebuilt part binds for it. */
const CANONICAL_PREFIX: ReadonlyMap<string, string> = new Map(
  Object.entries(OOXML_NAMESPACES).map(([prefix, { uri }]) => [uri, prefix]),
);

const XML_LOCAL_NAME = /^[\p{L}_][\p{L}\p{M}\p{N}._\-\u{00B7}\u{203F}-\u{2040}]*$/u;

export class InvalidPreservedAttributeError extends TaggedError("InvalidPreservedAttributeError")<{
  message: string;
  name: string;
}> {}

/** Whether an attribute remainder can be emitted under a bound canonical name. */
export const isSerializablePreservedAttribute = ({
  namespace,
  name,
}: Pick<PreservedAttribute, "namespace" | "name">): boolean =>
  name !== "xmlns" &&
  XML_LOCAL_NAME.test(name) &&
  (namespace === undefined || CANONICAL_PREFIX.has(namespace));

/**
 * The modelled set of an element whose serializer writes no attribute at all.
 *
 * `w:r`, `w:tr` and `w:sectPr` are rebuilt as bare start tags, so everything
 * their source wrote is the remainder. Named rather than spelled `new Set()`
 * at each call site, so the claim is visible and one place has to change when
 * one of them gains a modelled attribute.
 */
export const NO_MODELLED_ATTRIBUTES: ReadonlySet<string> = new Set();

type AttributeRemainderOptions = {
  /** The element, as the source wrote it. */
  element: XmlElement;
  /**
   * Local names this element's own parser reads.
   *
   * Local names rather than qualified ones, because folio's attribute readers
   * fall back to a local-name match across prefixes: an attribute the parser
   * would read under any binding has to be out of the remainder, or the save
   * writes the value twice under two spellings.
   */
  modelled: ReadonlySet<string>;
};

/**
 * Every attribute on the element the model has no field for.
 *
 * @returns the remainder, or `undefined` when the parser read them all — an
 *   empty list is never written, so a fully modelled element stays identical
 *   in the model.
 */
export const attributeRemainder = ({
  element,
  modelled,
}: AttributeRemainderOptions): PreservedAttribute[] | undefined => {
  const attributes = element.attributes;
  if (!attributes) {
    return undefined;
  }

  const remainder: PreservedAttribute[] = [];
  for (const [spelling, value] of Object.entries(attributes)) {
    if (value === undefined || spelling === "xmlns" || spelling.startsWith("xmlns:")) {
      continue;
    }
    const name = getLocalName(spelling);
    if (modelled.has(name)) {
      continue;
    }
    const namespace = resolveAttributeNamespaceUri(element, spelling);
    if (namespace !== undefined && !CANONICAL_PREFIX.has(namespace)) {
      continue;
    }
    remainder.push({
      ...(namespace === undefined ? {} : { namespace }),
      name,
      value: String(value),
    });
  }

  return remainder.length === 0 ? undefined : remainder;
};

const NAMES_BY_TABLE = new WeakMap<object, ReadonlySet<string>>();

/**
 * The remainder of an attribute bag, with the modelled set derived from the
 * record's own fields.
 *
 * `w:ind`, `w:spacing`, `w:framePr`, `w:tab`, a border side and `w:shd` are
 * bags rather than containers: the child dispatcher decides each of them
 * whole, so an attribute the record has no field for is kept only while the
 * reader takes nothing from the element at all. Passing the record's table
 * from `propertyElementAttributes.ts` rather than a set spelled at the call
 * site is what keeps the predicate derived from the model: a field added
 * without an attribute to name does not compile.
 */
export const readAttributeBag = (
  element: XmlElement,
  modelled: ModelledAttributes<string>,
): PreservedAttribute[] | undefined => {
  let names = NAMES_BY_TABLE.get(modelled);
  if (names === undefined) {
    names = new Set(modelledAttributeNames(modelled));
    NAMES_BY_TABLE.set(modelled, names);
  }
  return attributeRemainder({ element, modelled: names });
};

/** The name a fragment such as `w14:paraId="1F2E"` writes. */
const spelledName = (fragment: string): string => fragment.slice(0, fragment.indexOf("="));

/**
 * The element's attributes, modelled first and the remainder after them.
 *
 * @param modelled the fragments the serializer is about to emit, already
 *   spelled — they win a collision, because the model is the editable copy
 * @returns the fragments to write, so an element with neither writes nothing
 */
export const serializePreservedAttributes = (
  modelled: readonly string[],
  preserved: readonly PreservedAttribute[] | undefined,
): string[] => {
  if (preserved === undefined || preserved.length === 0) {
    return [...modelled];
  }

  const written = new Set(modelled.map(spelledName));
  const fragments = [...modelled];
  for (const attribute of preserved) {
    const { namespace, name, value } = attribute;
    if (!isSerializablePreservedAttribute(attribute)) {
      throw new InvalidPreservedAttributeError({
        message: "A preserved attribute must have an XML local name and a bound namespace",
        name,
      });
    }
    const prefix = namespace === undefined ? undefined : CANONICAL_PREFIX.get(namespace);
    const spelling = prefix === undefined ? name : `${prefix}:${name}`;
    if (written.has(spelling)) {
      continue;
    }
    written.add(spelling);
    fragments.push(`${spelling}="${escapeXmlAttribute(value)}"`);
  }
  return fragments;
};
