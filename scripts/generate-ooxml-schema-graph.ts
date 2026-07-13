import { createHash } from "node:crypto";
import { mkdir } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";
import { XMLParser } from "fast-xml-parser";
import JSZip from "jszip";

import sourceManifest from "../specifications/sources.json";

type XmlNode = Record<string, unknown>;

export type SchemaBundleInput = {
  documents: ReadonlyMap<string, string>;
  entrypoints: readonly string[];
  profile: "opc" | "transitional";
  sourceId: string;
  sourceSha256: string;
};

type SchemaDocument = {
  attributeFormDefault: "qualified" | "unqualified";
  elementFormDefault: "qualified" | "unqualified";
  node: XmlNode;
  path: string;
  prefixes: ReadonlyMap<string, string>;
  profile: "opc" | "transitional";
  sourceId: string;
  targetNamespace: string;
};

type SchemaGraphDocument = {
  imports: Array<{
    kind: "import" | "include";
    namespace?: string;
    schemaLocation?: string;
    target?: string;
  }>;
  namespace: string;
  path: string;
  profile: "opc" | "transitional";
  sourceId: string;
};

type SchemaGraphSymbol = {
  abstract?: boolean;
  base?: string;
  derivation?: "extension" | "restriction";
  document: string;
  enumValues?: string[];
  facets?: Array<{ kind: string; value: string }>;
  id: string;
  kind: DeclarationKind;
  memberTypes?: string[];
  mixed?: boolean;
  name: string;
  namespace: string;
  type?: string;
};

type SchemaGraphCompositor = {
  id: string;
  kind: "all" | "choice" | "sequence";
  maxOccurs: string;
  minOccurs: string;
  order: number;
  owner: string;
  parent?: string;
};

type SchemaGraphChild = {
  compositor?: string;
  enumValues?: string[];
  facets?: Array<{ kind: string; value: string }>;
  id: string;
  kind: "any" | "element" | "group";
  maxOccurs: string;
  minOccurs: string;
  name?: string;
  namespace?: string;
  order: number;
  owner: string;
  ref?: string;
  type?: string;
};

type SchemaGraphAttribute = {
  default?: string;
  enumValues?: string[];
  facets?: Array<{ kind: string; value: string }>;
  fixed?: string;
  id: string;
  kind: "any" | "attribute" | "group";
  name?: string;
  namespace?: string;
  order: number;
  owner: string;
  ref?: string;
  type?: string;
  use?: string;
};

export type OoxmlSchemaGraph = {
  attributes: SchemaGraphAttribute[];
  children: SchemaGraphChild[];
  compositors: SchemaGraphCompositor[];
  documents: SchemaGraphDocument[];
  inheritance: Array<{
    base: string;
    derived: string;
    method: "extension" | "restriction";
  }>;
  namespaces: Array<{
    documents: string[];
    uri: string;
  }>;
  profile: "docx-transitional";
  schemaVersion: 1;
  sources: Array<{
    id: string;
    sha256: string;
  }>;
  symbols: SchemaGraphSymbol[];
};

type DeclarationKind =
  | "attribute"
  | "attributeGroup"
  | "complexType"
  | "element"
  | "group"
  | "simpleType";

type GraphBuilder = {
  attributes: SchemaGraphAttribute[];
  attributeSequence: Map<string, number>;
  children: SchemaGraphChild[];
  childSequence: Map<string, number>;
  compositors: SchemaGraphCompositor[];
  compositorSequence: Map<string, number>;
  inheritance: OoxmlSchemaGraph["inheritance"];
  orderByParent: Map<string, number>;
  symbols: SchemaGraphSymbol[];
};

type WalkContext = {
  builder: GraphBuilder;
  document: SchemaDocument;
  owner: string;
  parentCompositor?: string;
};

class SchemaGraphError extends TaggedError("SchemaGraphError")<{
  message: string;
  cause?: unknown;
}>() {}

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(
  REPOSITORY_ROOT,
  "specifications",
  "generated",
  "docx-transitional-schema.gen.json",
);
const XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
const COMPOSITOR_KINDS = new Set(["all", "choice", "sequence"]);
const FACET_KINDS = new Set([
  "enumeration",
  "fractionDigits",
  "length",
  "maxExclusive",
  "maxInclusive",
  "maxLength",
  "minExclusive",
  "minInclusive",
  "minLength",
  "pattern",
  "totalDigits",
  "whiteSpace",
]);
const URL_SCHEME_RE = /^[a-z][a-z0-9+.-]*:/iu;
const xmlParser = new XMLParser({
  attributeNamePrefix: "@_",
  ignoreAttributes: false,
  parseAttributeValue: false,
  parseTagValue: false,
  preserveOrder: true,
  trimValues: true,
});

const compareStrings = (left: string, right: string): number => {
  if (left < right) {
    return -1;
  }
  if (left > right) {
    return 1;
  }
  return 0;
};

const compareById = (left: { id: string }, right: { id: string }): number =>
  compareStrings(left.id, right.id);

const nodeTag = (node: XmlNode): string | undefined =>
  Object.keys(node).find((key) => key !== ":@" && key !== "#text" && !key.startsWith("?"));

const localName = (qualifiedName: string): string =>
  qualifiedName.split(":").at(-1) ?? qualifiedName;

const nodeChildren = (node: XmlNode): XmlNode[] => {
  const tag = nodeTag(node);
  if (tag === undefined) {
    return [];
  }
  const value = node[tag];
  if (!Array.isArray(value)) {
    return [];
  }
  return value.filter((child): child is XmlNode => typeof child === "object" && child !== null);
};

const nodeAttributes = (node: XmlNode): Record<string, string> => {
  const raw = node[":@"];
  if (typeof raw !== "object" || raw === null || Array.isArray(raw)) {
    return {};
  }
  const attributes: Record<string, string> = {};
  for (const [key, value] of Object.entries(raw)) {
    if (key.startsWith("@_") && typeof value === "string") {
      attributes[key.slice(2)] = value;
    }
  }
  return attributes;
};

const findSchemaNode = (document: unknown, documentPath: string): XmlNode => {
  if (!Array.isArray(document)) {
    throw new SchemaGraphError({ message: `Invalid XML document: ${documentPath}` });
  }
  const schema = document.find(
    (node): node is XmlNode =>
      typeof node === "object" &&
      node !== null &&
      nodeTag(node) !== undefined &&
      localName(nodeTag(node) ?? "") === "schema",
  );
  if (schema === undefined) {
    throw new SchemaGraphError({ message: `Missing schema root: ${documentPath}` });
  }
  return schema;
};

const extractPrefixes = (attributes: Record<string, string>): ReadonlyMap<string, string> => {
  const prefixes = new Map<string, string>();
  for (const [name, value] of Object.entries(attributes)) {
    if (name === "xmlns") {
      prefixes.set("", value);
    } else if (name.startsWith("xmlns:")) {
      prefixes.set(name.slice("xmlns:".length), value);
    }
  }
  prefixes.set("xml", XML_NAMESPACE);
  return prefixes;
};

const resolveQName = (value: string, document: SchemaDocument): string => {
  const separator = value.indexOf(":");
  const prefix = separator === -1 ? "" : value.slice(0, separator);
  const name = separator === -1 ? value : value.slice(separator + 1);
  const namespace =
    document.prefixes.get(prefix) ?? (prefix === "" ? document.targetNamespace : undefined);
  if (namespace === undefined) {
    throw new SchemaGraphError({
      message: `Unknown namespace prefix "${prefix}" in ${document.path}`,
    });
  }
  return `{${namespace}}${name}`;
};

const symbolId = (kind: DeclarationKind, namespace: string, name: string): string =>
  `${kind}:{${namespace}}${name}`;

const nextSequence = (sequences: Map<string, number>, owner: string): number => {
  const next = (sequences.get(owner) ?? 0) + 1;
  sequences.set(owner, next);
  return next;
};

const nextOrder = (builder: GraphBuilder, parent: string): number => {
  const next = builder.orderByParent.get(parent) ?? 0;
  builder.orderByParent.set(parent, next + 1);
  return next;
};

const occurs = (value: string | undefined, fallback: string): string => value ?? fallback;

const localDeclarationNamespace = (
  attributes: Record<string, string>,
  document: SchemaDocument,
  declaration: "attribute" | "element",
): string => {
  const form = attributes["form"];
  if (form === "qualified") {
    return document.targetNamespace;
  }
  if (form === "unqualified") {
    return "";
  }
  if (declaration === "element") {
    return document.elementFormDefault === "qualified" ? document.targetNamespace : "";
  }
  return document.attributeFormDefault === "qualified" ? document.targetNamespace : "";
};

const readCompositorKind = (kind: string): "all" | "choice" | "sequence" => {
  if (kind === "all") {
    return "all";
  }
  if (kind === "choice") {
    return "choice";
  }
  return "sequence";
};

const readAttributeKind = (kind: string): "any" | "attribute" | "group" => {
  if (kind === "attribute") {
    return "attribute";
  }
  if (kind === "attributeGroup") {
    return "group";
  }
  return "any";
};

const collectSimpleType = (
  node: XmlNode,
  document: SchemaDocument,
): {
  base?: string;
  derivation?: "restriction";
  enumValues?: string[];
  facets?: Array<{ kind: string; value: string }>;
  memberTypes?: string[];
} => {
  const enumValues: string[] = [];
  const facets: Array<{ kind: string; value: string }> = [];
  let base: string | undefined;
  let derivation: "restriction" | undefined;
  let memberTypes: string[] | undefined;
  const visit = (current: XmlNode): void => {
    const tag = nodeTag(current);
    if (tag === undefined) {
      return;
    }
    const kind = localName(tag);
    const attributes = nodeAttributes(current);
    if (kind === "restriction" && attributes["base"] !== undefined) {
      base = resolveQName(attributes["base"], document);
      derivation = "restriction";
    }
    if (kind === "union" && attributes["memberTypes"] !== undefined) {
      memberTypes = attributes["memberTypes"]
        .split(/\s+/u)
        .filter((member) => member.length > 0)
        .map((member) => resolveQName(member, document));
    }
    const value = attributes["value"];
    if (value !== undefined && FACET_KINDS.has(kind)) {
      if (kind === "enumeration") {
        enumValues.push(value);
      } else {
        facets.push({ kind, value });
      }
    }
    for (const child of nodeChildren(current)) {
      visit(child);
    }
  };
  visit(node);
  return {
    ...(base === undefined ? {} : { base }),
    ...(derivation === undefined ? {} : { derivation }),
    ...(enumValues.length === 0 ? {} : { enumValues }),
    ...(facets.length === 0 ? {} : { facets }),
    ...(memberTypes === undefined ? {} : { memberTypes }),
  };
};

const addInheritance = (
  owner: string,
  method: "extension" | "restriction",
  base: string,
  builder: GraphBuilder,
): void => {
  if (!builder.inheritance.some((edge) => edge.derived === owner && edge.method === method)) {
    builder.inheritance.push({ base, derived: owner, method });
  }
};

const walkSchemaContent = (node: XmlNode, context: WalkContext): void => {
  const tag = nodeTag(node);
  if (tag === undefined) {
    return;
  }
  const kind = localName(tag);
  const attributes = nodeAttributes(node);

  if (kind === "extension" || kind === "restriction") {
    const base = attributes["base"];
    if (base !== undefined) {
      addInheritance(context.owner, kind, resolveQName(base, context.document), context.builder);
    }
    for (const child of nodeChildren(node)) {
      walkSchemaContent(child, context);
    }
    return;
  }

  if (COMPOSITOR_KINDS.has(kind)) {
    const sequence = nextSequence(context.builder.compositorSequence, context.owner);
    const id = `${context.owner}/compositor/${sequence}`;
    const parent = context.parentCompositor ?? context.owner;
    context.builder.compositors.push({
      id,
      kind: readCompositorKind(kind),
      maxOccurs: occurs(attributes["maxOccurs"], "1"),
      minOccurs: occurs(attributes["minOccurs"], "1"),
      order: nextOrder(context.builder, parent),
      owner: context.owner,
      ...(context.parentCompositor === undefined ? {} : { parent: context.parentCompositor }),
    });
    for (const child of nodeChildren(node)) {
      walkSchemaContent(child, { ...context, parentCompositor: id });
    }
    return;
  }

  if (kind === "element" || kind === "group" || kind === "any") {
    const sequence = nextSequence(context.builder.childSequence, context.owner);
    const id = `${context.owner}/child/${sequence}`;
    const parent = context.parentCompositor ?? context.owner;
    const ref = attributes["ref"];
    const name = attributes["name"];
    const child: SchemaGraphChild = {
      id,
      kind,
      maxOccurs: occurs(attributes["maxOccurs"], "1"),
      minOccurs: occurs(attributes["minOccurs"], "1"),
      order: nextOrder(context.builder, parent),
      owner: context.owner,
      ...(context.parentCompositor === undefined ? {} : { compositor: context.parentCompositor }),
      ...(name === undefined ? {} : { name }),
      ...(ref === undefined ? {} : { ref: resolveQName(ref, context.document) }),
      ...(attributes["type"] === undefined
        ? {}
        : { type: resolveQName(attributes["type"], context.document) }),
      ...(kind !== "element" || name === undefined
        ? {}
        : { namespace: localDeclarationNamespace(attributes, context.document, "element") }),
    };
    const inlineSimpleType = nodeChildren(node).find(
      (childNode) => localName(nodeTag(childNode) ?? "") === "simpleType",
    );
    if (inlineSimpleType !== undefined) {
      const metadata = collectSimpleType(inlineSimpleType, context.document);
      if (metadata.enumValues !== undefined) {
        child.enumValues = metadata.enumValues;
      }
      if (metadata.facets !== undefined) {
        child.facets = metadata.facets;
      }
    }
    context.builder.children.push(child);
    for (const childNode of nodeChildren(node)) {
      if (localName(nodeTag(childNode) ?? "") !== "simpleType") {
        walkSchemaContent(childNode, { ...context, owner: id, parentCompositor: undefined });
      }
    }
    return;
  }

  if (kind === "attribute" || kind === "attributeGroup" || kind === "anyAttribute") {
    const sequence = nextSequence(context.builder.attributeSequence, context.owner);
    const id = `${context.owner}/attribute/${sequence}`;
    const ref = attributes["ref"];
    const name = attributes["name"];
    const attribute: SchemaGraphAttribute = {
      id,
      kind: readAttributeKind(kind),
      order: sequence - 1,
      owner: context.owner,
      ...(name === undefined ? {} : { name }),
      ...(ref === undefined ? {} : { ref: resolveQName(ref, context.document) }),
      ...(attributes["type"] === undefined
        ? {}
        : { type: resolveQName(attributes["type"], context.document) }),
      ...(attributes["use"] === undefined ? {} : { use: attributes["use"] }),
      ...(attributes["default"] === undefined ? {} : { default: attributes["default"] }),
      ...(attributes["fixed"] === undefined ? {} : { fixed: attributes["fixed"] }),
      ...(kind !== "attribute" || name === undefined
        ? {}
        : { namespace: localDeclarationNamespace(attributes, context.document, "attribute") }),
    };
    const inlineSimpleType = nodeChildren(node).find(
      (childNode) => localName(nodeTag(childNode) ?? "") === "simpleType",
    );
    if (inlineSimpleType !== undefined) {
      const metadata = collectSimpleType(inlineSimpleType, context.document);
      if (metadata.enumValues !== undefined) {
        attribute.enumValues = metadata.enumValues;
      }
      if (metadata.facets !== undefined) {
        attribute.facets = metadata.facets;
      }
    }
    context.builder.attributes.push(attribute);
    return;
  }

  for (const child of nodeChildren(node)) {
    walkSchemaContent(child, context);
  }
};

const parseDocument = ({
  xml,
  documentPath,
  profile,
  sourceId,
}: {
  xml: string;
  documentPath: string;
  profile: "opc" | "transitional";
  sourceId: string;
}): SchemaDocument => {
  let parsed: unknown;
  try {
    parsed = xmlParser.parse(xml);
  } catch (cause) {
    throw new SchemaGraphError({ message: `Failed to parse ${documentPath}`, cause });
  }
  const node = findSchemaNode(parsed, documentPath);
  const attributes = nodeAttributes(node);
  const targetNamespace = attributes["targetNamespace"];
  if (targetNamespace === undefined) {
    throw new SchemaGraphError({ message: `Missing target namespace: ${documentPath}` });
  }
  const readFormDefault = (name: string): "qualified" | "unqualified" =>
    attributes[name] === "qualified" ? "qualified" : "unqualified";
  return {
    attributeFormDefault: readFormDefault("attributeFormDefault"),
    elementFormDefault: readFormDefault("elementFormDefault"),
    node,
    path: documentPath,
    prefixes: extractPrefixes(attributes),
    profile,
    sourceId,
    targetNamespace,
  };
};

const resolveImportTarget = (documentPath: string, schemaLocation: string): string => {
  const target = path.posix.normalize(
    path.posix.join(path.posix.dirname(documentPath), schemaLocation),
  );
  if (target === ".." || target.startsWith("../") || path.posix.isAbsolute(target)) {
    throw new SchemaGraphError({ message: `Schema import escapes its bundle: ${schemaLocation}` });
  }
  return target;
};

const readImports = (document: SchemaDocument): SchemaGraphDocument["imports"] => {
  const imports: SchemaGraphDocument["imports"] = [];
  for (const child of nodeChildren(document.node)) {
    const kind = localName(nodeTag(child) ?? "");
    if (kind !== "import" && kind !== "include") {
      continue;
    }
    const attributes = nodeAttributes(child);
    const schemaLocation = attributes["schemaLocation"];
    imports.push({
      kind,
      ...(attributes["namespace"] === undefined ? {} : { namespace: attributes["namespace"] }),
      ...(schemaLocation === undefined ? {} : { schemaLocation }),
      ...(schemaLocation === undefined || URL_SCHEME_RE.test(schemaLocation)
        ? {}
        : { target: resolveImportTarget(document.path, schemaLocation) }),
    });
  }
  return imports;
};

const loadBundleClosure = (bundle: SchemaBundleInput): SchemaDocument[] => {
  const documents: SchemaDocument[] = [];
  const loaded = new Set<string>();
  const queue = [...bundle.entrypoints].toSorted(compareStrings);
  while (queue.length > 0) {
    const documentPath = queue.shift();
    if (documentPath === undefined || loaded.has(documentPath)) {
      continue;
    }
    const xml = bundle.documents.get(documentPath);
    if (xml === undefined) {
      throw new SchemaGraphError({ message: `Missing imported schema: ${documentPath}` });
    }
    const document = parseDocument({
      xml,
      documentPath,
      profile: bundle.profile,
      sourceId: bundle.sourceId,
    });
    loaded.add(documentPath);
    documents.push(document);
    for (const edge of readImports(document)) {
      if (edge.target !== undefined && !loaded.has(edge.target)) {
        queue.push(edge.target);
      }
    }
    queue.sort(compareStrings);
  }
  return documents;
};

const addDocumentDeclarations = (document: SchemaDocument, builder: GraphBuilder): void => {
  for (const child of nodeChildren(document.node)) {
    const kind = localName(nodeTag(child) ?? "");
    const declarationKind = readDeclarationKind(kind);
    if (declarationKind === undefined) {
      continue;
    }
    const attributes = nodeAttributes(child);
    const name = attributes["name"];
    if (name === undefined) {
      continue;
    }
    const id = symbolId(declarationKind, document.targetNamespace, name);
    const symbol: SchemaGraphSymbol = {
      document: document.path,
      id,
      kind: declarationKind,
      name,
      namespace: document.targetNamespace,
      ...(attributes["abstract"] === "true" || attributes["abstract"] === "1"
        ? { abstract: true }
        : {}),
      ...(attributes["mixed"] === "true" || attributes["mixed"] === "1" ? { mixed: true } : {}),
      ...(attributes["type"] === undefined
        ? {}
        : { type: resolveQName(attributes["type"], document) }),
    };
    if (declarationKind === "simpleType") {
      Object.assign(symbol, collectSimpleType(child, document));
    }
    builder.symbols.push(symbol);
    for (const content of nodeChildren(child)) {
      walkSchemaContent(content, { builder, document, owner: id });
    }
  }
};

const readDeclarationKind = (kind: string): DeclarationKind | undefined => {
  if (kind === "attribute") {
    return "attribute";
  }
  if (kind === "attributeGroup") {
    return "attributeGroup";
  }
  if (kind === "complexType") {
    return "complexType";
  }
  if (kind === "element") {
    return "element";
  }
  if (kind === "group") {
    return "group";
  }
  if (kind === "simpleType") {
    return "simpleType";
  }
  return undefined;
};

const createBuilder = (): GraphBuilder => ({
  attributes: [],
  attributeSequence: new Map(),
  children: [],
  childSequence: new Map(),
  compositors: [],
  compositorSequence: new Map(),
  inheritance: [],
  orderByParent: new Map(),
  symbols: [],
});

export const buildOoxmlSchemaGraph = (bundles: readonly SchemaBundleInput[]): OoxmlSchemaGraph => {
  const documents = bundles
    .flatMap(loadBundleClosure)
    .toSorted((left, right) =>
      compareStrings(`${left.profile}:${left.path}`, `${right.profile}:${right.path}`),
    );
  const builder = createBuilder();
  for (const document of documents) {
    addDocumentDeclarations(document, builder);
  }
  const documentsByNamespace = new Map<string, string[]>();
  for (const document of documents) {
    const paths = documentsByNamespace.get(document.targetNamespace) ?? [];
    paths.push(document.path);
    documentsByNamespace.set(document.targetNamespace, paths);
  }
  return {
    attributes: builder.attributes.toSorted(compareById),
    children: builder.children.toSorted(compareById),
    compositors: builder.compositors.toSorted(compareById),
    documents: documents.map((document) => ({
      imports: readImports(document),
      namespace: document.targetNamespace,
      path: document.path,
      profile: document.profile,
      sourceId: document.sourceId,
    })),
    inheritance: builder.inheritance.toSorted((left, right) =>
      compareStrings(
        `${left.derived}:${left.method}:${left.base}`,
        `${right.derived}:${right.method}:${right.base}`,
      ),
    ),
    namespaces: [...documentsByNamespace.entries()]
      .map(([uri, paths]) => ({ documents: paths.toSorted(compareStrings), uri }))
      .toSorted((left, right) => compareStrings(left.uri, right.uri)),
    profile: "docx-transitional",
    schemaVersion: 1,
    sources: bundles
      .map(({ sourceId: id, sourceSha256: sha256 }) => ({ id, sha256 }))
      .toSorted((left, right) => compareStrings(left.id, right.id)),
    symbols: builder.symbols.toSorted(compareById),
  };
};

const sha256 = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex");

const readPinnedArtifact = async (
  sourceId: string,
): Promise<{ bytes: Uint8Array; sha256: string }> => {
  const source = sourceManifest.sources.find(({ id }) => id === sourceId);
  if (source === undefined || source.sourceType !== "artifact") {
    throw new SchemaGraphError({ message: `Missing artifact source: ${sourceId}` });
  }
  const cachePath = path.resolve(REPOSITORY_ROOT, source.cachePath);
  const file = Bun.file(cachePath);
  if (!(await file.exists())) {
    throw new SchemaGraphError({
      message: `Missing cached source ${sourceId}; run bun run specifications:fetch`,
    });
  }
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (bytes.length !== source.bytes || sha256(bytes) !== source.sha256) {
    throw new SchemaGraphError({ message: `Cached source does not match manifest: ${sourceId}` });
  }
  return { bytes, sha256: source.sha256 };
};

const readNestedXsdBundle = async ({
  outerBytes,
  innerPath,
}: {
  outerBytes: Uint8Array;
  innerPath: string;
}): Promise<ReadonlyMap<string, string>> => {
  const outer = await JSZip.loadAsync(outerBytes);
  const innerEntry = outer.file(innerPath);
  if (innerEntry === null) {
    throw new SchemaGraphError({ message: `Missing nested schema archive: ${innerPath}` });
  }
  const inner = await JSZip.loadAsync(await innerEntry.async("uint8array"));
  const documents = new Map<string, string>();
  for (const entry of Object.values(inner.files).toSorted((left, right) =>
    compareStrings(left.name, right.name),
  )) {
    if (entry.dir || !entry.name.endsWith(".xsd")) {
      continue;
    }
    // oxlint-disable-next-line no-await-in-loop -- archive order is deterministic and bounded
    documents.set(entry.name, await entry.async("text"));
  }
  return documents;
};

const buildPinnedGraph = async (): Promise<OoxmlSchemaGraph> => {
  const [transitional, opc] = await Promise.all([
    readPinnedArtifact("ecma-376-part-4-transitional"),
    readPinnedArtifact("ecma-376-part-2-opc"),
  ]);
  const [transitionalDocuments, opcDocuments] = await Promise.all([
    readNestedXsdBundle({
      outerBytes: transitional.bytes,
      innerPath: "OfficeOpenXML-XMLSchema-Transitional.zip",
    }),
    readNestedXsdBundle({
      outerBytes: opc.bytes,
      innerPath: "OpenPackagingConventions-XMLSchema.zip",
    }),
  ]);
  return buildOoxmlSchemaGraph([
    {
      documents: opcDocuments,
      entrypoints: [...opcDocuments.keys()].toSorted(compareStrings),
      profile: "opc",
      sourceId: "ecma-376-part-2-opc",
      sourceSha256: opc.sha256,
    },
    {
      documents: transitionalDocuments,
      entrypoints: ["wml.xsd"],
      profile: "transitional",
      sourceId: "ecma-376-part-4-transitional",
      sourceSha256: transitional.sha256,
    },
  ]);
};

const serializeGraph = (graph: OoxmlSchemaGraph): string => `${JSON.stringify(graph, null, 2)}\n`;

const main = async (): Promise<void> => {
  const command = process.argv.at(2) ?? "write";
  if (command !== "check" && command !== "write") {
    throw new SchemaGraphError({
      message: "Usage: bun scripts/generate-ooxml-schema-graph.ts [check|write]",
    });
  }
  const generated = serializeGraph(await buildPinnedGraph());
  if (command === "check") {
    const current = await Bun.file(OUTPUT_PATH).text();
    if (current !== generated) {
      throw new SchemaGraphError({
        message: "Generated schema graph is stale; run bun run specifications:generate",
      });
    }
    process.stdout.write("Specification schema graph is current\n");
    return;
  }
  await mkdir(path.dirname(OUTPUT_PATH), { recursive: true });
  await Bun.write(OUTPUT_PATH, generated);
  process.stdout.write(`Wrote ${path.relative(REPOSITORY_ROOT, OUTPUT_PATH)}\n`);
};

if (import.meta.main) {
  main().catch((cause: unknown) => {
    const message = cause instanceof Error ? cause.message : String(cause);
    process.stderr.write(`${message}\n`);
    process.exitCode = 1;
  });
}
