/**
 * The space of container/child and element/attribute pairs folio must not lose.
 *
 * Everything downstream — the survival law, the container contract's keys, the
 * coverage check — reads its universe from here, so there is one derivation of
 * "what the schema allows inside the parts folio rebuilds" and no hand-listed
 * mirror of it. The scoping is deliberately the same as
 * `scripts/generate-strict-value-encodings.ts`: the roots of the rebuilt parts
 * and the namespaces a WordprocessingML part can carry inline. A slot outside
 * that scope reaches a package as a part folio copies through, so it cannot be
 * dropped by a parser that never reads it.
 */

import path from "node:path";

import type { OoxmlSchemaGraph } from "../../generate-ooxml-schema-graph";

export const WML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** @see INLINE_NAMESPACES in `scripts/generate-strict-value-encodings.ts`. */
export const INLINE_NAMESPACES: ReadonlySet<string> = new Set([
  WML_NAMESPACE,
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://schemas.openxmlformats.org/drawingml/2006/picture",
  "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas",
  "http://schemas.openxmlformats.org/officeDocument/2006/math",
]);

/** @see REBUILT_PART_ROOTS in `scripts/generate-strict-value-encodings.ts`. */
export const REBUILT_PART_ROOTS: readonly string[] = [
  "comments",
  "document",
  "endnotes",
  "footnotes",
  "ftr",
  "hdr",
];

const GRAPH_PATH = path.join(
  path.resolve(import.meta.dir, "../../.."),
  "specifications/generated/docx-transitional-schema.gen.json",
);

export type QualifiedName = { namespace: string; name: string };

/** `{namespace}local` — the spelling the schema graph uses for a reference. */
export const qualify = ({ namespace, name }: QualifiedName): string => `{${namespace}}${name}`;

/**
 * A container is an element *declared with a particular complex type*.
 *
 * WordprocessingML reuses a name across types — `w:tblPr` is `CT_TblPr` on a
 * table and `CT_TblPrBase` inside a style — and the two allow different
 * children, so the type belongs in the identity.
 */
export type ContainerId = { element: QualifiedName; typeQName: string };

export const containerKey = ({ element, typeQName }: ContainerId): string =>
  `${qualify(element)}|${typeQName}`;

export type ChildSlot = {
  container: ContainerId;
  child: QualifiedName;
  /** The complex type the child is declared with, when the schema gives it one. */
  childTypeQName: string | undefined;
  minOccurs: string;
  maxOccurs: string;
  /** Position among the container's flattened particles; the validator's order check reads it. */
  order: number;
};

export type AttributeSlot = {
  container: ContainerId;
  attribute: QualifiedName;
  typeQName: string | undefined;
  required: boolean;
  default: string | undefined;
  fixed: string | undefined;
};

export const childSlotKey = (slot: ChildSlot): string =>
  `${containerKey(slot.container)}/${qualify(slot.child)}`;

export const attributeSlotKey = (slot: AttributeSlot): string =>
  `${containerKey(slot.container)}@${qualify(slot.attribute)}`;

type Index = {
  attributesByOwner: ReadonlyMap<string, OoxmlSchemaGraph["attributes"]>;
  baseOf: ReadonlyMap<string, string>;
  byId: ReadonlyMap<string, OoxmlSchemaGraph["symbols"][number]>;
  childrenByOwner: ReadonlyMap<string, OoxmlSchemaGraph["children"]>;
  compositorKinds: ReadonlyMap<string, string>;
  /** Compositors that need no member: `minOccurs="0"` on themselves or on an ancestor. */
  optionalCompositors: ReadonlySet<string>;
  globalAttributes: ReadonlyMap<string, OoxmlSchemaGraph["symbols"][number]>;
  graph: OoxmlSchemaGraph;
};

/**
 * Compositors a document may leave empty.
 *
 * A particle's own `minOccurs` is not the whole answer: the particles inside
 * `EG_ContentRowContent` all declare `minOccurs="1"`, and the group reference
 * that pulls them into `CT_Tbl` declares `minOccurs="0"`. Reading the particle
 * alone makes every run-level element look required inside a table, which is
 * how a generated fixture ends up several kilobytes of markup nobody asked for.
 */
const optionalCompositorsOf = (graph: OoxmlSchemaGraph): Set<string> => {
  const byId = new Map(graph.compositors.map((compositor) => [compositor.id, compositor]));
  const optional = new Set<string>();
  const resolve = (id: string, seen: Set<string>): boolean => {
    if (optional.has(id)) {
      return true;
    }
    const compositor = byId.get(id);
    if (compositor === undefined || seen.has(id)) {
      return false;
    }
    seen.add(id);
    const isOptional =
      compositor.minOccurs === "0" ||
      (compositor.parent !== undefined && resolve(compositor.parent, seen));
    if (isOptional) {
      optional.add(id);
    }
    return isOptional;
  };
  for (const { id } of graph.compositors) {
    resolve(id, new Set());
  }
  return optional;
};

const groupBy = <T extends { owner: string }>(items: readonly T[]): Map<string, T[]> => {
  const grouped = new Map<string, T[]>();
  for (const item of items) {
    const list = grouped.get(item.owner);
    if (list) {
      list.push(item);
      continue;
    }
    grouped.set(item.owner, [item]);
  }
  return grouped;
};

export const buildIndex = (graph: OoxmlSchemaGraph): Index => ({
  attributesByOwner: groupBy(graph.attributes),
  baseOf: new Map(graph.inheritance.map(({ derived, base }) => [derived, base])),
  byId: new Map(graph.symbols.map((symbol) => [symbol.id, symbol])),
  childrenByOwner: groupBy(graph.children),
  compositorKinds: new Map(graph.compositors.map(({ id, kind }) => [id, kind])),
  optionalCompositors: optionalCompositorsOf(graph),
  globalAttributes: new Map(
    graph.symbols
      .filter((symbol) => symbol.kind === "attribute")
      .map((symbol) => [qualify(symbol), symbol]),
  ),
  graph,
});

export type SchemaIndex = Index;

type ResolvedAttribute = {
  attribute: QualifiedName;
  typeQName: string | undefined;
  required: boolean;
  default: string | undefined;
  fixed: string | undefined;
  enumValues: readonly string[] | undefined;
};

/**
 * The attributes a complex type declares, following attribute groups and its base.
 *
 * The graph already resolves `attributeFormDefault`, so a local declaration
 * carries `""` when the document spells it unqualified (all of DrawingML) and
 * its target namespace when it does not (WordprocessingML). An empty namespace
 * is data here, not a missing value: it is how an attribute with no prefix is
 * spelled.
 */
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
    if (attribute.kind !== "attribute") {
      continue;
    }
    if (attribute.ref) {
      const global = index.globalAttributes.get(attribute.ref);
      if (global) {
        resolved.push({
          attribute: { namespace: global.namespace, name: global.name },
          typeQName: global.type,
          required: attribute.use === "required",
          default: attribute.default,
          fixed: attribute.fixed,
          enumValues: attribute.enumValues,
        });
      }
      continue;
    }
    if (attribute.name === undefined) {
      continue;
    }
    resolved.push({
      attribute: { namespace: attribute.namespace ?? "", name: attribute.name },
      typeQName: attribute.type,
      required: attribute.use === "required",
      default: attribute.default,
      fixed: attribute.fixed,
      enumValues: attribute.enumValues,
    });
  }

  const base = index.baseOf.get(ownerId);
  const baseType = base === undefined ? undefined : index.byId.get(`complexType:${base}`);
  if (baseType) {
    resolved.push(...attributesOf(index, baseType.id, seen));
  }
  return resolved;
};

export type ResolvedChild = {
  child: QualifiedName;
  typeQName: string | undefined;
  minOccurs: string;
  maxOccurs: string;
  /** The compositor the particle sits in, so a required-sibling walk can tell a choice from a sequence. */
  compositorId: string | undefined;
};

/**
 * The element particles a complex type allows, in declaration order.
 *
 * Model groups are inlined where they are referenced, which is what the
 * validator's ordinal check does too, so the index of a particle here is the
 * ordinal a document has to respect.
 */
export const childrenOf = (
  index: Index,
  ownerId: string,
  seen = new Set<string>(),
  inherited: { optional: boolean } = { optional: false },
): ResolvedChild[] => {
  if (seen.has(ownerId)) {
    return [];
  }
  seen.add(ownerId);

  // An extension contributes the base's particles *before* its own, which is
  // also how `corpus-schema-validator.ts` assigns ordinals. Emitting a
  // `w:tblGridChange` ahead of the `w:gridCol` it follows would be reported as
  // an out-of-order child, and the fixture would be blamed for a loss it did
  // not cause.
  const base = index.baseOf.get(ownerId);
  const baseType = base === undefined ? undefined : index.byId.get(`complexType:${base}`);
  const resolved: ResolvedChild[] =
    baseType === undefined ? [] : childrenOf(index, baseType.id, seen, inherited);

  for (const child of index.childrenByOwner.get(ownerId) ?? []) {
    if (child.kind === "group") {
      if (child.ref) {
        resolved.push(
          ...childrenOf(index, `group:${child.ref}`, seen, {
            optional:
              inherited.optional ||
              child.minOccurs === "0" ||
              index.optionalCompositors.has(child.compositor ?? ""),
          }),
        );
      }
      continue;
    }
    if (child.kind !== "element") {
      continue;
    }
    const declared = child.ref ? index.byId.get(`element:${child.ref}`) : undefined;
    const namespace = declared?.namespace ?? child.namespace;
    const name = declared?.name ?? child.name;
    if (namespace === undefined || name === undefined) {
      continue;
    }
    const optional =
      inherited.optional ||
      child.minOccurs === "0" ||
      index.optionalCompositors.has(child.compositor ?? "");
    resolved.push({
      child: { namespace, name },
      typeQName: declared?.type ?? child.type,
      minOccurs: optional ? "0" : child.minOccurs,
      maxOccurs: child.maxOccurs,
      compositorId: child.compositor,
    });
  }

  return resolved;
};

export type Container = {
  id: ContainerId;
  children: ChildSlot[];
  attributes: AttributeSlot[];
  /**
   * The cheapest chain of elements from a rebuilt part's root to this container.
   *
   * The first step is the root element itself. A container reachable from more
   * than one root keeps the chain from `w:document`, and among chains of equal
   * length the one that avoids {@link DETOUR_ELEMENTS}, so the census is stable
   * across runs and its fixtures look like documents.
   */
  path: readonly ContainerId[];
};

export type ContainerSpace = {
  containers: ReadonlyMap<string, Container>;
  index: Index;
};

/**
 * Wrappers a document may nest anywhere, which make a shortest path unrealistic.
 *
 * `m:deg` is two steps from `w:body` through `<w:ins><m:rad>` and four through
 * `<w:p><w:r>`, because a tracked-change wrapper accepts paragraph content
 * directly. The short chain is schema-valid and no document is written that
 * way, so a fixture built along it measures how folio treats a bare `w:ins` in
 * a body rather than how it treats the pair. Costing a detour more than a step
 * keeps the chain on the structural spine without hand-listing the spine.
 *
 * These are the elements that are *transparent*: they wrap content of their
 * parent's own kind rather than introducing a kind of their own.
 */
const DETOUR_ELEMENTS: ReadonlySet<string> = new Set([
  "bdo",
  "customXml",
  "del",
  "dir",
  "fldSimple",
  "hyperlink",
  "ins",
  "moveFrom",
  "moveTo",
  "sdt",
  "sdtContent",
  "smartTag",
  "subDoc",
]);

/**
 * Containers whose content is block-level, and the blocks they actually hold.
 *
 * The schema lets a `w:body` hold an `m:oMath` directly; a document puts the
 * equation in a paragraph. Charging the direct edge a detour routes the fixture
 * through `w:p` without saying anything about what folio should do with the
 * direct form — that stays a pair of its own, tested where it is declared.
 */
const BLOCK_CONTAINERS: ReadonlySet<string> = new Set([
  "body",
  "comment",
  "endnote",
  "footnote",
  "ftr",
  "hdr",
  "tc",
  "txbxContent",
]);

const BLOCK_CHILDREN: ReadonlySet<string> = new Set(["p", "sdt", "tbl"]);

const DETOUR_COST = 6;

/**
 * The two penalties stack, and they have to.
 *
 * A run-level `w:ins` is reachable from `w:body` in one step, because the
 * schema lets a tracked-change wrapper hold run content anywhere paragraph
 * content is allowed. That chain builds a `<w:body><w:ins><m:acc/></w:ins>` —
 * a tracked insertion with no paragraph — which folio discards, so every pair
 * inside `CT_RunTrackChange` would read as lost. Charging both the wrapper and
 * the unusual edge out of the block container routes the chain through `w:p`,
 * where tracked changes are the marks folio actually models.
 */
const stepCost = (parent: QualifiedName, child: QualifiedName): number => {
  let cost = 1;
  if (child.namespace === WML_NAMESPACE && DETOUR_ELEMENTS.has(child.name)) {
    cost += DETOUR_COST;
  }
  if (
    parent.namespace === WML_NAMESPACE &&
    BLOCK_CONTAINERS.has(parent.name) &&
    !(child.namespace === WML_NAMESPACE && BLOCK_CHILDREN.has(child.name))
  ) {
    cost += DETOUR_COST;
  }
  return cost;
};

/**
 * Every container reachable from a rebuilt part's root, with the cheapest path to it.
 *
 * The walk is Dijkstra over the element graph: minimal paths keep the fixtures
 * small, and the detour cost keeps them shaped like documents. A deep fixture
 * costs more to build and has more ways to be invalid for reasons that have
 * nothing to do with the pair under test.
 */
export const buildContainerSpace = (graph: OoxmlSchemaGraph): ContainerSpace => {
  const index = buildIndex(graph);
  const containers = new Map<string, Container>();

  const best = new Map<string, { cost: number; path: readonly ContainerId[] }>();
  const queue: Array<{ id: ContainerId; path: readonly ContainerId[]; cost: number }> = [];

  /**
   * Every root is a source, but `w:document` is free and the others are not.
   *
   * A single unweighted walk would hand `w:tbl` the one-step chain from
   * `w:ftr`, and the fixture builder would then have to synthesise a header
   * part for a container that sits in every document. The penalty is larger
   * than any path inside the body, so a container the body can reach always
   * keeps the body's chain, and one it cannot still gets the chain it has.
   */
  const ROOT_PENALTY = 1000;

  const seedRoot = (root: string): void => {
    const element = index.byId.get(`element:{${WML_NAMESPACE}}${root}`);
    if (element?.type === undefined) {
      return;
    }
    const id: ContainerId = {
      element: { namespace: WML_NAMESPACE, name: root },
      typeQName: element.type,
    };
    const cost = root === "document" ? 0 : ROOT_PENALTY;
    best.set(containerKey(id), { cost, path: [id] });
    queue.push({ id, path: [id], cost });
  };

  /** A small graph and a small queue, so the frontier is scanned rather than heaped. */
  const drain = (): void => {
    while (queue.length > 0) {
      let cheapest = 0;
      for (let candidate = 1; candidate < queue.length; candidate += 1) {
        // SAFETY: both indices are inside the queue the loop bounds.
        if (
          (queue[candidate] as { cost: number }).cost < (queue[cheapest] as { cost: number }).cost
        ) {
          cheapest = candidate;
        }
      }
      // SAFETY: `cheapest` indexes a queue the loop condition proves non-empty.
      const entry = queue.splice(cheapest, 1)[0] as {
        id: ContainerId;
        path: readonly ContainerId[];
        cost: number;
      };
      if (best.get(containerKey(entry.id))?.cost !== entry.cost) {
        continue;
      }
      visitContainer(entry);
    }
  };

  const visitContainer = ({
    id,
    path: chain,
    cost,
  }: {
    id: ContainerId;
    path: readonly ContainerId[];
    cost: number;
  }): void => {
    const key = containerKey(id);
    const resolved = childrenOf(index, `complexType:${id.typeQName}`);

    const children: ChildSlot[] = resolved.map((child, order) => ({
      container: id,
      child: child.child,
      childTypeQName: child.typeQName,
      minOccurs: child.minOccurs,
      maxOccurs: child.maxOccurs,
      order,
    }));
    const attributes: AttributeSlot[] = attributesOf(index, `complexType:${id.typeQName}`).map(
      (attribute) => ({
        container: id,
        attribute: attribute.attribute,
        typeQName: attribute.typeQName,
        required: attribute.required,
        default: attribute.default,
        fixed: attribute.fixed,
      }),
    );
    containers.set(key, { id, children, attributes, path: chain });

    for (const child of resolved) {
      if (child.typeQName === undefined || !INLINE_NAMESPACES.has(child.child.namespace)) {
        continue;
      }
      const childId: ContainerId = { element: child.child, typeQName: child.typeQName };
      const childKey = containerKey(childId);
      const childCost = cost + stepCost(id.element, child.child);
      const known = best.get(childKey);
      if (known !== undefined && known.cost <= childCost) {
        continue;
      }
      const childPath = [...chain, childId];
      best.set(childKey, { cost: childCost, path: childPath });
      queue.push({ id: childId, path: childPath, cost: childCost });
    }
  };

  for (const root of REBUILT_PART_ROOTS) {
    seedRoot(root);
  }
  drain();
  for (const [key, container] of containers) {
    const resolved = best.get(key);
    if (resolved !== undefined) {
      containers.set(key, { ...container, path: resolved.path });
    }
  }

  return { containers, index };
};

/** Every child slot in the space, in a stable order. */
export const childSlots = (space: ContainerSpace): ChildSlot[] =>
  [...space.containers.values()]
    .flatMap((container) => container.children)
    .filter((slot) => INLINE_NAMESPACES.has(slot.child.namespace))
    .sort((left, right) => childSlotKey(left).localeCompare(childSlotKey(right)));

/** Every attribute slot in the space, in a stable order. */
export const attributeSlots = (space: ContainerSpace): AttributeSlot[] =>
  [...space.containers.values()]
    .flatMap((container) => container.attributes)
    .sort((left, right) => attributeSlotKey(left).localeCompare(attributeSlotKey(right)));

let graphPromise: Promise<OoxmlSchemaGraph> | undefined;

/** Memoised per process; the graph is ~2.7 MB and every caller reads the same one. */
export const loadContainerSchemaGraph = (): Promise<OoxmlSchemaGraph> => {
  graphPromise ??= Bun.file(GRAPH_PATH).json() as Promise<OoxmlSchemaGraph>;
  return graphPromise;
};

let spacePromise: Promise<ContainerSpace> | undefined;

export const loadContainerSpace = (): Promise<ContainerSpace> => {
  spacePromise ??= loadContainerSchemaGraph().then(buildContainerSpace);
  return spacePromise;
};
