/**
 * Check an OOXML part against the committed schema graph, in signature form.
 *
 * The corpus gate groups thousands of files into a handful of defects, so every
 * verdict here must be a function of the schema alone: two files that break the
 * same declaration produce byte-identical violations, with no path, no index,
 * no attribute value and no count of that document's elements. A violation that
 * carried the offending value would split one defect into as many signatures as
 * the corpus has spellings of it.
 *
 * The check is deliberately partial. Word writes packages a strict validator
 * rejects, and folio must round-trip them, so a false positive here costs a
 * real investigation while a miss costs nothing: `xsd:any` admits everything,
 * markup-compatibility and foreign namespaces are skipped whole, a union whose
 * member is a built-in accepts every value, and child order is only checked
 * where the content model is a single flat sequence of distinct names. What is
 * left is the class of defect folio can actually cause: an element or attribute
 * that no declaration allows.
 */

import path from "node:path";

import { Result, TaggedError } from "better-result";
import { XMLParser } from "fast-xml-parser";

import type { OoxmlSchemaGraph } from "../generate-ooxml-schema-graph";
import { orderedParticlesByOwner } from "./ooxml-schema-graph";

export const SCHEMA_VIOLATION_KINDS = {
  /** Element not allowed as a child of its parent's type. */
  unknownElement: "unknown-element",
  unknownAttribute: "unknown-attribute",
  missingRequiredAttribute: "missing-required-attribute",
  badEnumValue: "bad-enum-value",
  /** Only for pure-sequence content models. */
  outOfOrderChild: "out-of-order-child",
  /** The part's root element is not a global element in the graph. */
  unknownRoot: "unknown-root",
} as const;

export type SchemaViolationKind =
  (typeof SCHEMA_VIOLATION_KINDS)[keyof typeof SCHEMA_VIOLATION_KINDS];

export type SchemaViolation = {
  kind: SchemaViolationKind;
  /** Slash path of local names from the root, e.g. "document/body/p/pPr/rPr". No indices. */
  path: string;
  /** Qualified name that broke, e.g. "{<ns>}rFonts" or an attribute name. */
  name: string;
  /** Short human detail, no per-file particulars. */
  detail: string;
};

export type ValidatePartOptions = {
  graph: OoxmlSchemaGraph;
  /** The part's XML text. */
  xml: string;
  /** Stop after this many violations; the caller only needs signatures. */
  limit?: number;
};

type SchemaSymbol = OoxmlSchemaGraph["symbols"][number];
type SchemaChild = OoxmlSchemaGraph["children"][number];
type SchemaAttribute = OoxmlSchemaGraph["attributes"][number];
type SchemaCompositor = OoxmlSchemaGraph["compositors"][number];
type SchemaInheritance = OoxmlSchemaGraph["inheritance"][number];

class SchemaPartParseError extends TaggedError("SchemaPartParseError")<{
  message: string;
  cause?: unknown;
}> {}

const GRAPH_PATH = path.resolve(
  import.meta.dir,
  "..",
  "..",
  "specifications",
  "generated",
  "docx-transitional-schema.gen.json",
);

const DEFAULT_VIOLATION_LIMIT = 50;

const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE = "http://www.w3.org/2000/xmlns/";
const XML_SCHEMA_INSTANCE_NAMESPACE = "http://www.w3.org/2001/XMLSchema-instance";
const MARKUP_COMPATIBILITY_NAMESPACE =
  "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** Namespaces whose attributes describe the markup, not the document model. */
const IGNORED_ATTRIBUTE_NAMESPACES: ReadonlySet<string> = new Set([
  XMLNS_NAMESPACE,
  XML_SCHEMA_INSTANCE_NAMESPACE,
  MARKUP_COMPATIBILITY_NAMESPACE,
]);

const XMLNS_DECLARATION = "xmlns";
const XMLNS_PREFIX_DECLARATION = "xmlns:";
const XML_PREFIX = "xml";
const DEFAULT_PREFIX = "";

/** fast-xml-parser's `preserveOrder` shape: the attribute bag and the text leaf. */
const ATTRIBUTES_KEY = ":@";
const NON_ELEMENT_KEY_PREFIXES = ["#", "?"];

const SYMBOL_PREFIXES = {
  attribute: "attribute:",
  attributeGroup: "attributeGroup:",
  complexType: "complexType:",
  element: "element:",
  group: "group:",
  simpleType: "simpleType:",
} as const;

const PARTICLE_KINDS = {
  any: "any",
  element: "element",
  group: "group",
} as const;

const SEQUENCE_COMPOSITOR = "sequence";
const REQUIRED_USE = "required";
const EXTENSION_DERIVATION = "extension";

const UNKNOWN_ROOT_DETAIL = "root element is not a global element declaration";
const OUT_OF_ORDER_DETAIL = "child precedes a sibling the sequence declares earlier";

const PARSER_OPTIONS = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: "",
  allowBooleanAttributes: true,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
};

type ParsedNode = Record<string, unknown>;
type ParsedElement = { node: ParsedNode; tag: string };

/** Prefix to namespace URI, with `""` for the default declaration. */
type NamespaceScope = ReadonlyMap<string, string>;

/** The `xml` prefix is bound by the XML specification, not by a declaration. */
const ROOT_SCOPE: NamespaceScope = new Map([[XML_PREFIX, XML_NAMESPACE]]);

type ResolvedChild = { qname: string; typeQName: string | undefined };
type ResolvedAttribute = {
  qname: string;
  typeQName: string | undefined;
  required: boolean;
};

type ContentModel = {
  children: ResolvedChild[];
  allowsAny: boolean;
  /** The particles form one flat sequence, so declaration order is meaningful. */
  ordered: boolean;
};

type AttributeModel = { attributes: ResolvedAttribute[]; allowsAny: boolean };

type TypeModel = {
  children: ResolvedChild[];
  allowsAnyChild: boolean;
  ordered: boolean;
  attributes: ResolvedAttribute[];
  allowsAnyAttribute: boolean;
};

type ResolvedType = {
  childByQName: ReadonlyMap<string, ResolvedChild>;
  allowsAnyChild: boolean;
  attributeByQName: ReadonlyMap<string, ResolvedAttribute>;
  requiredAttributes: readonly ResolvedAttribute[];
  allowsAnyAttribute: boolean;
  /** Present only when the order check is admissible; see `contentModelFor`. */
  ordinalByQName: ReadonlyMap<string, number> | undefined;
};

type SchemaIndex = {
  symbolsById: ReadonlyMap<string, SchemaSymbol>;
  globalElementByQName: ReadonlyMap<string, SchemaSymbol>;
  childrenByOwner: ReadonlyMap<string, readonly SchemaChild[]>;
  attributesByOwner: ReadonlyMap<string, readonly SchemaAttribute[]>;
  compositorsById: ReadonlyMap<string, SchemaCompositor>;
  compositorCountByOwner: ReadonlyMap<string, number>;
  inheritanceByDerived: ReadonlyMap<string, SchemaInheritance>;
  knownNamespaces: ReadonlySet<string>;
  typeCache: Map<string, ResolvedType>;
  enumerationCache: Map<string, ReadonlySet<string> | undefined>;
};

const qualify = (namespace: string, local: string): string =>
  namespace === "" ? local : `{${namespace}}${local}`;

/** The namespace a qualified name carries, or the empty string when unqualified. */
const namespaceOfQName = (qname: string): string =>
  qname.startsWith("{") ? qname.slice(1, qname.indexOf("}")) : "";

const splitName = (name: string): { prefix: string; local: string } => {
  const colon = name.indexOf(":");
  if (colon === -1) {
    return { prefix: DEFAULT_PREFIX, local: name };
  }
  return { prefix: name.slice(0, colon), local: name.slice(colon + 1) };
};

const buildIndex = (graph: OoxmlSchemaGraph): SchemaIndex => {
  const symbolsById = new Map(graph.symbols.map((symbol) => [symbol.id, symbol]));
  const globalElementByQName = new Map(
    graph.symbols
      .filter((symbol) => symbol.kind === "element")
      .map((symbol) => [qualify(symbol.namespace, symbol.name), symbol]),
  );
  // The ordinal this walk scores a document's children against is the one the
  // survival census builds its fixtures at, so both read the same derivation.
  const childrenByOwner = orderedParticlesByOwner(graph.children);
  const attributesByOwner = orderedParticlesByOwner(graph.attributes);
  const compositorCountByOwner = new Map<string, number>();
  for (const compositor of graph.compositors) {
    compositorCountByOwner.set(
      compositor.owner,
      (compositorCountByOwner.get(compositor.owner) ?? 0) + 1,
    );
  }
  return {
    symbolsById,
    globalElementByQName,
    childrenByOwner,
    attributesByOwner,
    compositorsById: new Map(graph.compositors.map((compositor) => [compositor.id, compositor])),
    compositorCountByOwner,
    inheritanceByDerived: new Map(graph.inheritance.map((edge) => [edge.derived, edge])),
    knownNamespaces: new Set(graph.namespaces.map(({ uri }) => uri)),
    typeCache: new Map(),
    enumerationCache: new Map(),
  };
};

const indexByGraph = new WeakMap<OoxmlSchemaGraph, SchemaIndex>();

const indexFor = (graph: OoxmlSchemaGraph): SchemaIndex => {
  const cached = indexByGraph.get(graph);
  if (cached !== undefined) {
    return cached;
  }
  const built = buildIndex(graph);
  indexByGraph.set(graph, built);
  return built;
};

const elementParticle = (index: SchemaIndex, particle: SchemaChild): ResolvedChild => {
  if (particle.ref !== undefined) {
    return { qname: particle.ref, typeQName: index.globalElementByQName.get(particle.ref)?.type };
  }
  return {
    qname: qualify(particle.namespace ?? "", particle.name ?? ""),
    typeQName: particle.type,
  };
};

const attributeParticle = (index: SchemaIndex, particle: SchemaAttribute): ResolvedAttribute => {
  const required = particle.use === REQUIRED_USE;
  if (particle.ref !== undefined) {
    const global = index.symbolsById.get(`${SYMBOL_PREFIXES.attribute}${particle.ref}`);
    return { qname: particle.ref, typeQName: global?.type, required };
  }
  return {
    qname: qualify(particle.namespace ?? "", particle.name ?? ""),
    typeQName: particle.type,
    required,
  };
};

/**
 * Order is only decidable where the model cannot reorder itself.
 *
 * A `choice` or `all` compositor, a nested compositor, a wildcard, or a name
 * reachable through two particles all make "this child came too early" a guess,
 * and a guess here reads as a folio defect in the corpus census. So the walk
 * demands a single top-level `sequence` of element particles, and the caller
 * additionally demands that every name in it be distinct. `maxOccurs` is not
 * consulted: within one flat sequence the ordinals stay monotone whether a
 * particle repeats or not, and it is only a repeated *name* that is ambiguous.
 */
const contentModelFor = (
  index: SchemaIndex,
  ownerId: string,
  expanding: Set<string>,
): ContentModel => {
  if (expanding.has(ownerId)) {
    return { children: [], allowsAny: false, ordered: false };
  }
  expanding.add(ownerId);
  const particles = index.childrenByOwner.get(ownerId) ?? [];
  const children: ResolvedChild[] = [];
  let allowsAny = false;
  let ordered = (index.compositorCountByOwner.get(ownerId) ?? 0) <= 1;
  for (const particle of particles) {
    const compositor =
      particle.compositor === undefined
        ? undefined
        : index.compositorsById.get(particle.compositor);
    if (compositor !== undefined) {
      ordered &&= compositor.kind === SEQUENCE_COMPOSITOR && compositor.parent === undefined;
    }
    if (particle.kind === PARTICLE_KINDS.any) {
      allowsAny = true;
      ordered = false;
      continue;
    }
    if (particle.kind === PARTICLE_KINDS.group) {
      if (particle.ref === undefined) {
        ordered = false;
        continue;
      }
      const expanded = contentModelFor(index, `${SYMBOL_PREFIXES.group}${particle.ref}`, expanding);
      children.push(...expanded.children);
      allowsAny ||= expanded.allowsAny;
      ordered &&= expanded.ordered;
      continue;
    }
    children.push(elementParticle(index, particle));
  }
  expanding.delete(ownerId);
  return { children, allowsAny, ordered };
};

const attributeModelFor = (
  index: SchemaIndex,
  ownerId: string,
  expanding: Set<string>,
): AttributeModel => {
  if (expanding.has(ownerId)) {
    return { attributes: [], allowsAny: false };
  }
  expanding.add(ownerId);
  const particles = index.attributesByOwner.get(ownerId) ?? [];
  const attributes: ResolvedAttribute[] = [];
  let allowsAny = false;
  for (const particle of particles) {
    if (particle.kind === PARTICLE_KINDS.any) {
      allowsAny = true;
      continue;
    }
    if (particle.kind === PARTICLE_KINDS.group) {
      if (particle.ref === undefined) {
        continue;
      }
      const expanded = attributeModelFor(
        index,
        `${SYMBOL_PREFIXES.attributeGroup}${particle.ref}`,
        expanding,
      );
      attributes.push(...expanded.attributes);
      allowsAny ||= expanded.allowsAny;
      continue;
    }
    attributes.push(attributeParticle(index, particle));
  }
  expanding.delete(ownerId);
  return { attributes, allowsAny };
};

const EMPTY_TYPE_MODEL: TypeModel = {
  children: [],
  allowsAnyChild: false,
  ordered: true,
  attributes: [],
  allowsAnyAttribute: false,
};

/**
 * An extension contributes the base's particles and attributes before its own;
 * a restriction replaces them. The committed graph carries no complexType
 * restriction, so the branch below only ever takes the extension path, but a
 * future schema revision must not silently inherit through a restriction.
 */
const typeModelFor = (index: SchemaIndex, typeQName: string, deriving: Set<string>): TypeModel => {
  const typeId = `${SYMBOL_PREFIXES.complexType}${typeQName}`;
  if (deriving.has(typeId)) {
    return EMPTY_TYPE_MODEL;
  }
  deriving.add(typeId);
  const inherited = index.inheritanceByDerived.get(typeId);
  const base =
    inherited !== undefined && inherited.method === EXTENSION_DERIVATION
      ? typeModelFor(index, inherited.base, deriving)
      : EMPTY_TYPE_MODEL;
  const content = contentModelFor(index, typeId, new Set());
  const attributeModel = attributeModelFor(index, typeId, new Set());
  deriving.delete(typeId);
  return {
    children: [...base.children, ...content.children],
    allowsAnyChild: base.allowsAnyChild || content.allowsAny,
    ordered: base.ordered && content.ordered,
    attributes: [...base.attributes, ...attributeModel.attributes],
    allowsAnyAttribute: base.allowsAnyAttribute || attributeModel.allowsAny,
  };
};

const resolveType = (index: SchemaIndex, typeQName: string): ResolvedType => {
  const cached = index.typeCache.get(typeQName);
  if (cached !== undefined) {
    return cached;
  }
  const model = typeModelFor(index, typeQName, new Set());
  const childByQName = new Map(model.children.map((child) => [child.qname, child]));
  const attributeByQName = new Map(
    model.attributes.map((attribute) => [attribute.qname, attribute]),
  );
  const orderIsDecidable = model.ordered && childByQName.size === model.children.length;
  const resolved: ResolvedType = {
    childByQName,
    allowsAnyChild: model.allowsAnyChild,
    attributeByQName,
    requiredAttributes: [...attributeByQName.values()].filter(({ required }) => required),
    allowsAnyAttribute: model.allowsAnyAttribute,
    ordinalByQName: orderIsDecidable
      ? new Map(model.children.map((child, ordinal) => [child.qname, ordinal]))
      : undefined,
  };
  index.typeCache.set(typeQName, resolved);
  return resolved;
};

/**
 * The values a simpleType admits, or `undefined` for "anything".
 *
 * `undefined` covers a built-in (the graph carries no `xsd:*` symbol) and any
 * union with a built-in member, which is how ECMA-376's `ST_OnOff` reaches
 * `true|false|1|0` alongside `on|off`: its members are `xsd:boolean` and
 * `ST_OnOff1`. Widening there beats enumerating boolean lexical forms by hand.
 */
const enumerationOf = (
  index: SchemaIndex,
  typeQName: string,
  following: Set<string>,
): ReadonlySet<string> | undefined => {
  const typeId = `${SYMBOL_PREFIXES.simpleType}${typeQName}`;
  if (following.has(typeId)) {
    return undefined;
  }
  const symbol = index.symbolsById.get(typeId);
  if (symbol === undefined) {
    return undefined;
  }
  if (symbol.enumValues !== undefined) {
    return new Set(symbol.enumValues);
  }
  following.add(typeId);
  const resolved = resolveDerivedEnumeration(index, symbol, following);
  following.delete(typeId);
  return resolved;
};

const resolveDerivedEnumeration = (
  index: SchemaIndex,
  symbol: SchemaSymbol,
  following: Set<string>,
): ReadonlySet<string> | undefined => {
  if (symbol.memberTypes !== undefined) {
    const union = new Set<string>();
    for (const member of symbol.memberTypes) {
      const values = enumerationOf(index, member, following);
      if (values === undefined) {
        return undefined;
      }
      for (const value of values) {
        union.add(value);
      }
    }
    return union;
  }
  if (symbol.base === undefined) {
    return undefined;
  }
  return enumerationOf(index, symbol.base, following);
};

const enumerationFor = (index: SchemaIndex, typeQName: string): ReadonlySet<string> | undefined => {
  if (index.enumerationCache.has(typeQName)) {
    return index.enumerationCache.get(typeQName);
  }
  const values = enumerationOf(index, typeQName, new Set());
  index.enumerationCache.set(typeQName, values);
  return values;
};

const attributesOf = (node: ParsedNode): Record<string, unknown> | undefined => {
  const bag = node[ATTRIBUTES_KEY];
  return typeof bag === "object" && bag !== null ? (bag as Record<string, unknown>) : undefined;
};

const tagOf = (node: ParsedNode): string | undefined => {
  for (const key of Object.keys(node)) {
    if (key === ATTRIBUTES_KEY) {
      continue;
    }
    if (NON_ELEMENT_KEY_PREFIXES.some((prefix) => key.startsWith(prefix))) {
      continue;
    }
    return key;
  }
  return undefined;
};

const elementsOf = (nodes: readonly unknown[]): ParsedElement[] => {
  const elements: ParsedElement[] = [];
  for (const candidate of nodes) {
    if (typeof candidate !== "object" || candidate === null) {
      continue;
    }
    const node = candidate as ParsedNode;
    const tag = tagOf(node);
    if (tag === undefined) {
      continue;
    }
    elements.push({ node, tag });
  }
  return elements;
};

const childrenOf = ({ node, tag }: ParsedElement): ParsedElement[] => {
  const nested = node[tag];
  return Array.isArray(nested) ? elementsOf(nested) : [];
};

const extendScope = (parent: NamespaceScope, node: ParsedNode): NamespaceScope => {
  const attributes = attributesOf(node);
  if (attributes === undefined) {
    return parent;
  }
  let scope: Map<string, string> | undefined;
  for (const [name, value] of Object.entries(attributes)) {
    if (name === XMLNS_DECLARATION) {
      scope ??= new Map(parent);
      scope.set(DEFAULT_PREFIX, String(value));
      continue;
    }
    if (!name.startsWith(XMLNS_PREFIX_DECLARATION)) {
      continue;
    }
    scope ??= new Map(parent);
    scope.set(name.slice(XMLNS_PREFIX_DECLARATION.length), String(value));
  }
  return scope ?? parent;
};

/** An unprefixed element takes the default namespace; an unprefixed attribute takes none. */
const elementNamespace = (prefix: string, scope: NamespaceScope): string | undefined =>
  prefix === DEFAULT_PREFIX ? (scope.get(DEFAULT_PREFIX) ?? "") : scope.get(prefix);

const attributeNamespace = (prefix: string, scope: NamespaceScope): string | undefined =>
  prefix === DEFAULT_PREFIX ? "" : scope.get(prefix);

const isValidatableNamespace = (index: SchemaIndex, namespace: string): boolean =>
  namespace !== MARKUP_COMPATIBILITY_NAMESPACE && index.knownNamespaces.has(namespace);

type WalkState = {
  index: SchemaIndex;
  violations: SchemaViolation[];
  limit: number;
};

const isFull = (state: WalkState): boolean => state.violations.length >= state.limit;

const record = (state: WalkState, violation: SchemaViolation): void => {
  if (isFull(state)) {
    return;
  }
  state.violations.push(violation);
};

const observedAttributes = (node: ParsedNode, scope: NamespaceScope): Map<string, string> => {
  const observed = new Map<string, string>();
  const attributes = attributesOf(node);
  if (attributes === undefined) {
    return observed;
  }
  for (const [name, value] of Object.entries(attributes)) {
    if (name === XMLNS_DECLARATION || name.startsWith(XMLNS_PREFIX_DECLARATION)) {
      continue;
    }
    const { prefix, local } = splitName(name);
    const namespace = attributeNamespace(prefix, scope);
    // An undeclared prefix is a well-formedness fault, not a schema one.
    if (namespace === undefined || IGNORED_ATTRIBUTE_NAMESPACES.has(namespace)) {
      continue;
    }
    observed.set(qualify(namespace, local), String(value));
  }
  return observed;
};

/**
 * An attribute in a namespace this schema does not describe.
 *
 * Producers decorate WordprocessingML elements with extension attributes and
 * declare the prefix ignorable: `w14:paraId` on `w:p` is what every recent Word
 * writes. Word accepts them, so folio must, and a validator that called them
 * schema violations would report the producer rather than folio on almost every
 * real package.
 */
const isForeignAttribute = (index: SchemaIndex, qname: string): boolean => {
  const namespace = namespaceOfQName(qname);
  return namespace !== "" && !index.knownNamespaces.has(namespace);
};

const checkAttributes = (
  state: WalkState,
  element: ParsedElement,
  scope: NamespaceScope,
  type: ResolvedType,
  typeQName: string,
  elementPath: string,
): void => {
  const observed = observedAttributes(element.node, scope);
  for (const [qname, value] of observed) {
    if (isFull(state)) {
      return;
    }
    const declared = type.attributeByQName.get(qname);
    if (declared === undefined) {
      if (type.allowsAnyAttribute || isForeignAttribute(state.index, qname)) {
        continue;
      }
      record(state, {
        kind: SCHEMA_VIOLATION_KINDS.unknownAttribute,
        path: elementPath,
        name: qname,
        detail: `attribute not declared on ${typeQName}`,
      });
      continue;
    }
    if (declared.typeQName === undefined) {
      continue;
    }
    const admitted = enumerationFor(state.index, declared.typeQName);
    if (admitted === undefined || admitted.has(value)) {
      continue;
    }
    record(state, {
      kind: SCHEMA_VIOLATION_KINDS.badEnumValue,
      path: elementPath,
      name: qname,
      detail: `value outside the enumeration of ${declared.typeQName}`,
    });
  }
  for (const required of type.requiredAttributes) {
    if (isFull(state)) {
      return;
    }
    if (observed.has(required.qname)) {
      continue;
    }
    record(state, {
      kind: SCHEMA_VIOLATION_KINDS.missingRequiredAttribute,
      path: elementPath,
      name: required.qname,
      detail: `required attribute missing on ${typeQName}`,
    });
  }
};

const walkElement = (
  state: WalkState,
  element: ParsedElement,
  scope: NamespaceScope,
  typeQName: string,
  elementPath: string,
): void => {
  if (isFull(state)) {
    return;
  }
  const type = resolveType(state.index, typeQName);
  checkAttributes(state, element, scope, type, typeQName, elementPath);
  let highestOrdinal = -1;
  let orderReported = false;
  for (const child of childrenOf(element)) {
    if (isFull(state)) {
      return;
    }
    const childScope = extendScope(scope, child.node);
    const { prefix, local } = splitName(child.tag);
    const namespace = elementNamespace(prefix, childScope);
    if (namespace === undefined || !isValidatableNamespace(state.index, namespace)) {
      continue;
    }
    const qname = qualify(namespace, local);
    const childPath = `${elementPath}/${local}`;
    const particle = type.childByQName.get(qname);
    if (particle === undefined) {
      if (!type.allowsAnyChild) {
        record(state, {
          kind: SCHEMA_VIOLATION_KINDS.unknownElement,
          path: childPath,
          name: qname,
          detail: `element not declared in the content model of ${typeQName}`,
        });
      }
      // No declared type, so the subtree has nothing to be checked against.
      continue;
    }
    const ordinal = type.ordinalByQName?.get(qname);
    if (ordinal !== undefined && !orderReported && ordinal < highestOrdinal) {
      orderReported = true;
      record(state, {
        kind: SCHEMA_VIOLATION_KINDS.outOfOrderChild,
        path: childPath,
        name: qname,
        detail: OUT_OF_ORDER_DETAIL,
      });
    }
    if (ordinal !== undefined && ordinal > highestOrdinal) {
      highestOrdinal = ordinal;
    }
    if (particle.typeQName === undefined) {
      continue;
    }
    walkElement(state, child, childScope, particle.typeQName, childPath);
  }
};

export const validateOoxmlPart = ({
  graph,
  xml,
  limit = DEFAULT_VIOLATION_LIMIT,
}: ValidatePartOptions): SchemaViolation[] => {
  const parsed = Result.try({
    try: (): unknown => new XMLParser(PARSER_OPTIONS).parse(xml),
    catch: (cause) => new SchemaPartParseError({ message: "part is not well-formed XML", cause }),
  });
  // Well-formedness is the corpus classifier's verdict (`malformed-document-xml`);
  // a part that never parsed has no schema signature to report here.
  if (parsed.isErr() || !Array.isArray(parsed.value)) {
    return [];
  }
  const index = indexFor(graph);
  const root = elementsOf(parsed.value).at(0);
  if (root === undefined) {
    return [];
  }
  const scope = extendScope(ROOT_SCOPE, root.node);
  const { prefix, local } = splitName(root.tag);
  const namespace = elementNamespace(prefix, scope);
  // A part rooted in foreign markup is not folio's to validate.
  if (namespace === undefined || !isValidatableNamespace(index, namespace)) {
    return [];
  }
  const qname = qualify(namespace, local);
  const declaration = index.globalElementByQName.get(qname);
  if (declaration?.type === undefined) {
    return [
      {
        kind: SCHEMA_VIOLATION_KINDS.unknownRoot,
        path: local,
        name: qname,
        detail: UNKNOWN_ROOT_DETAIL,
      },
    ];
  }
  const state: WalkState = { index, violations: [], limit };
  walkElement(state, root, scope, declaration.type, local);
  return state.violations;
};

/**
 * Whether the validator has an order to score a type's children against.
 *
 * `contentModelFor` refuses one wherever the model can reorder itself — a
 * `choice`, an `all`, a wildcard, a name two particles reach — so a caller
 * asking "did the order check run?" cannot read the answer off a clean verdict:
 * silence means both "in order" and "no order to be in". The census's
 * sequence-row order test needs the two apart, and this is the validator's own
 * answer rather than a second derivation beside it.
 */
export const ordersChildrenOf = (graph: OoxmlSchemaGraph, typeQName: string): boolean =>
  resolveType(indexFor(graph), typeQName).ordinalByQName !== undefined;

let graphPromise: Promise<OoxmlSchemaGraph> | undefined;

/** Memoised per process: the graph is ~2.7 MB and every part reads the same one. */
export const loadSchemaGraph = (): Promise<OoxmlSchemaGraph> => {
  graphPromise ??= Bun.file(GRAPH_PATH).json() as Promise<OoxmlSchemaGraph>;
  return graphPromise;
};
