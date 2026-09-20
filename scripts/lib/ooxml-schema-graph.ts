/**
 * Walking the committed OOXML schema graph.
 *
 * `specifications/generated/docx-transitional-schema.gen.json` is the only
 * offline record of what a Transitional package may declare, and two checks
 * derive slot sets from it: `generate-strict-value-encodings.ts`, which finds
 * the attributes a Strict producer spells with a unit suffix, and
 * `check-reserved-value-coverage.ts`, which finds the slots that carry a
 * reserved value. They must agree on which slots a rebuilt part can reach, so
 * the reachability walk lives here rather than in either of them.
 */

import { readFile } from "node:fs/promises";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "../..");

export const SCHEMA_GRAPH_PATH = path.join(
  REPO_ROOT,
  "specifications/generated/docx-transitional-schema.gen.json",
);

export const WML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Namespaces whose content a WordprocessingML part can carry inline.
 *
 * Chart, diagram and chart-drawing markup reaches a document only as a
 * reference to a part of its own, and folio copies those parts through
 * unchanged, so their slots would never be consulted.
 */
export const INLINE_NAMESPACES: ReadonlySet<string> = new Set([
  WML_NAMESPACE,
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://schemas.openxmlformats.org/drawingml/2006/picture",
  "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
  "http://schemas.openxmlformats.org/officeDocument/2006/math",
]);

/**
 * Roots of the parts folio rebuilds from the model.
 *
 * WordprocessingML reuses element names across parts — `w:w` is character
 * scaling in a run and a frameset splitbar width in `webSettings.xml` — so the
 * slot table is restricted to what a rebuilt part's root can actually contain.
 * The other inline vocabularies reach a document only through a
 * `a:graphicData` payload the schema types as `xs:any`, which no reachability
 * walk can follow, so they contribute every slot they declare.
 */
export const REBUILT_PART_ROOTS: readonly string[] = [
  "comments",
  "document",
  "endnotes",
  "footnotes",
  "ftr",
  "hdr",
];

export type SchemaSymbol = {
  base?: string;
  enumValues?: string[];
  facets?: Array<{ kind: string; value: string }>;
  id: string;
  kind: string;
  memberTypes?: string[];
  name: string;
  namespace: string;
  type?: string;
};

export type AttributeDeclaration = {
  default?: string;
  fixed?: string;
  kind: string;
  name?: string;
  /** Position in the owner's attribute list, as declared. */
  order?: number;
  owner: string;
  ref?: string;
  type?: string;
  use?: string;
};

export type ChildDeclaration = {
  kind: string;
  name?: string;
  namespace?: string;
  /** Position in the owner's content model; the order a serializer must write. */
  order?: number;
  owner: string;
  ref?: string;
  type?: string;
};

export type SchemaGraph = {
  attributes: AttributeDeclaration[];
  children: ChildDeclaration[];
  inheritance: Array<{ base: string; derived: string }>;
  namespaces: Array<{ uri: string }>;
  symbols: SchemaSymbol[];
};

export const loadSchemaGraph = async (graphPath = SCHEMA_GRAPH_PATH): Promise<SchemaGraph> =>
  JSON.parse(await readFile(graphPath, "utf8")) as SchemaGraph;

export const buildIndex = (graph: SchemaGraph) => {
  const byId = new Map(graph.symbols.map((symbol) => [symbol.id, symbol]));
  const attributesByOwner = new Map<string, AttributeDeclaration[]>();
  for (const attribute of graph.attributes) {
    const list = attributesByOwner.get(attribute.owner);
    if (list) {
      list.push(attribute);
      continue;
    }
    attributesByOwner.set(attribute.owner, [attribute]);
  }
  const baseOf = new Map(graph.inheritance.map(({ derived, base }) => [derived, base]));
  const globalAttributes = new Map(
    graph.symbols
      .filter((symbol) => symbol.kind === "attribute")
      .map((symbol) => [`{${symbol.namespace}}${symbol.name}`, symbol]),
  );
  return { attributesByOwner, baseOf, byId, globalAttributes };
};

export type Index = ReturnType<typeof buildIndex>;

export type ResolvedAttribute = { name: string; type: string; default?: string; fixed?: string };

/** Attributes a complex type declares, following attribute groups and its base. */
export const attributesOf = (
  index: Index,
  ownerId: string,
  seen = new Set<string>(),
): ResolvedAttribute[] => {
  if (seen.has(ownerId)) {
    return [];
  }
  seen.add(ownerId);

  const resolved: ResolvedAttribute[] = [];
  for (const attribute of index.attributesByOwner.get(ownerId) ?? []) {
    if (attribute.kind === "group") {
      if (attribute.ref) {
        resolved.push(...attributesOf(index, `attributeGroup:${attribute.ref}`, seen));
      }
      continue;
    }
    if (attribute.ref) {
      const global = index.globalAttributes.get(attribute.ref);
      if (global?.type) {
        resolved.push({ name: global.name, type: global.type });
      }
      continue;
    }
    if (attribute.name && attribute.type) {
      resolved.push({
        name: attribute.name,
        type: attribute.type,
        ...(attribute.default === undefined ? {} : { default: attribute.default }),
        ...(attribute.fixed === undefined ? {} : { fixed: attribute.fixed }),
      });
    }
  }

  const base = index.baseOf.get(ownerId);
  const baseType = base === undefined ? undefined : index.byId.get(`complexType:${base}`);
  if (baseType) {
    resolved.push(...attributesOf(index, baseType.id, seen));
  }
  return resolved;
};

/** WordprocessingML elements a rebuilt part's root can contain, with the types they take. */
export const reachableWmlElementTypes = (
  graph: SchemaGraph,
  index: Index,
  roots: readonly string[] = REBUILT_PART_ROOTS,
): ReadonlyMap<string, ReadonlySet<string>> => {
  const childrenByOwner = new Map<string, ChildDeclaration[]>();
  for (const child of graph.children) {
    const list = childrenByOwner.get(child.owner ?? "");
    if (list) {
      list.push(child);
      continue;
    }
    childrenByOwner.set(child.owner ?? "", [child]);
  }

  const reached = new Map<string, Set<string>>();
  const reach = (name: string, type: string | undefined): void => {
    const types = reached.get(name);
    if (types === undefined) {
      reached.set(name, new Set(type === undefined ? [] : [type]));
      return;
    }
    if (type !== undefined) {
      types.add(type);
    }
  };
  const visited = new Set<string>();
  const visit = (ownerId: string): void => {
    if (visited.has(ownerId)) {
      return;
    }
    visited.add(ownerId);
    for (const child of childrenByOwner.get(ownerId) ?? []) {
      if (child.kind === "group") {
        if (child.ref) {
          visit(`group:${child.ref}`);
        }
        continue;
      }
      if (child.kind !== "element") {
        continue;
      }
      const declared = child.ref ? index.byId.get(`element:${child.ref}`) : undefined;
      const namespace = declared?.namespace ?? child.namespace;
      const name = declared?.name ?? child.name;
      const type = declared?.type ?? child.type;
      if (name !== undefined && namespace === WML_NAMESPACE) {
        reach(name, type);
      }
      if (type !== undefined) {
        visit(`complexType:${type}`);
      }
    }
    const base = index.baseOf.get(ownerId);
    if (base !== undefined) {
      visit(`complexType:${base}`);
    }
  };

  for (const root of roots) {
    const element = index.byId.get(`element:{${WML_NAMESPACE}}${root}`);
    reach(root, element?.type);
    if (element?.type) {
      visit(`complexType:${element.type}`);
    }
  }
  return reached;
};

/** Every complex type an element of a given qualified name can be declared with. */
export const elementTypes = (
  graph: SchemaGraph,
  index: Index,
  roots: readonly string[] = REBUILT_PART_ROOTS,
): Map<string, Set<string>> => {
  const reachable = reachableWmlElementTypes(graph, index, roots);
  const types = new Map<string, Set<string>>();
  const record = (namespace: string, name: string, type: string | undefined): void => {
    if (type === undefined || !INLINE_NAMESPACES.has(namespace)) {
      return;
    }
    if (namespace === WML_NAMESPACE && reachable.get(name)?.has(type) !== true) {
      return;
    }
    const key = `${namespace} ${name}`;
    const set = types.get(key);
    if (set) {
      set.add(type);
      return;
    }
    types.set(key, new Set([type]));
  };

  const byId = new Map(graph.symbols.map((symbol) => [symbol.id, symbol]));
  for (const symbol of graph.symbols) {
    if (symbol.kind === "element") {
      record(symbol.namespace, symbol.name, symbol.type);
    }
  }
  for (const child of graph.children) {
    if (child.kind !== "element") {
      continue;
    }
    if (child.ref) {
      const element = byId.get(`element:${child.ref}`);
      if (element) {
        record(element.namespace, element.name, element.type);
      }
      continue;
    }
    if (child.name && child.namespace !== undefined) {
      record(child.namespace, child.name, child.type);
    }
  }
  return types;
};

export const localName = (qualifiedName: string): string =>
  qualifiedName.slice(qualifiedName.indexOf("}") + 1);

/** The slot key both checks use: `"<namespace URI> <element> @<attribute>"`. */
export const slotKey = (namespace: string, element: string, attribute: string | null): string =>
  `${namespace} ${element}${attribute === null ? "" : ` @${attribute}`}`;
