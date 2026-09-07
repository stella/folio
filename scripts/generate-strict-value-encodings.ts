/**
 * Generate the table that says how a Strict-spelled value is written in Transitional.
 *
 * ECMA-376 gives many attributes two spellings of the same value: a plain
 * number in the unit the attribute is defined in, and a string carrying its own
 * unit (`155.85pt`, `20%`). Strict producers write the string form; Transitional
 * consumers want the number. folio rebuilds every package as Transitional, so
 * content captured verbatim from a Strict source has to be re-spelled, and the
 * set of slots that needs it is far too large to hand-list: a couple of hundred
 * across WordprocessingML, DrawingML, and the math and picture vocabularies.
 *
 * The set is therefore derived from the Transitional schema graph: every
 * attribute whose type is a union of a numeric member and a measure- or
 * percentage-patterned member. The unit each union encodes is not in the schema
 * (nothing there says a twip is a twentieth of a point), so it comes from
 * {@link UNION_ENCODINGS}, which the generator requires to be total over the
 * unions it finds: a schema refresh that adds one fails here instead of
 * silently dropping its slots.
 *
 * Usage:
 *   bun scripts/generate-strict-value-encodings.ts write
 *   bun scripts/generate-strict-value-encodings.ts check
 */

import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const GRAPH_PATH = path.join(
  REPO_ROOT,
  "specifications/generated/docx-transitional-schema.gen.json",
);
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/core/src/docx/strictValueEncodings.gen.ts");

class GenerateStrictValueEncodingsError extends TaggedError("GenerateStrictValueEncodingsError")<{
  message: string;
}> {}

const WML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * Namespaces whose content a WordprocessingML part can carry inline.
 *
 * Chart, diagram and chart-drawing markup reaches a document only as a
 * reference to a part of its own, and folio copies those parts through
 * unchanged, so their slots would never be consulted.
 */
const INLINE_NAMESPACES: ReadonlySet<string> = new Set([
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
const REBUILT_PART_ROOTS: readonly string[] = [
  "comments",
  "document",
  "endnotes",
  "footnotes",
  "ftr",
  "hdr",
];

/**
 * How each dual-spelling union writes its value as a number.
 *
 * `null` excludes a union whose numeric member has no single unit: in
 * `ST_DecimalNumberOrPercent` the number means whole percent under
 * `w:zoom/@w:percent` and half-points under
 * `w:readModeInkLockDown/@w:fontSz`, and every element carrying it lives in
 * `settings.xml`, which folio copies through rather than rebuilds.
 */
const UNION_ENCODINGS: Readonly<Record<string, { measure?: string; percent?: string } | null>> = {
  ST_AdjCoordinate: { measure: "emu" },
  ST_Coordinate: { measure: "emu" },
  ST_Coordinate32: { measure: "emu" },
  ST_DecimalNumberOrPercent: null,
  ST_FixedPercentage: { percent: "thousandthPercent" },
  ST_HpsMeasure: { measure: "halfPoints" },
  ST_MeasurementOrPercent: { measure: "twips", percent: "fiftiethPercent" },
  ST_Percentage: { percent: "thousandthPercent" },
  ST_PositiveFixedPercentage: { percent: "thousandthPercent" },
  ST_PositivePercentage: { percent: "thousandthPercent" },
  ST_SignedHpsMeasure: { measure: "halfPoints" },
  ST_SignedTwipsMeasure: { measure: "twips" },
  ST_TextFontScalePercentOrPercentString: { percent: "thousandthPercent" },
  ST_TextPoint: { measure: "hundredthPoints" },
  ST_TextScale: { percent: "wholePercent" },
  ST_TextSpacingPercentOrPercentString: { percent: "thousandthPercent" },
  ST_TwipsMeasure: { measure: "twips" },
};

type SchemaSymbol = {
  base?: string;
  facets?: Array<{ kind: string; value: string }>;
  id: string;
  kind: string;
  memberTypes?: string[];
  name: string;
  namespace: string;
  type?: string;
};

type AttributeDeclaration = {
  kind: string;
  name?: string;
  owner: string;
  ref?: string;
  type?: string;
};

type ChildDeclaration = {
  kind: string;
  name?: string;
  namespace?: string;
  owner: string;
  ref?: string;
  type?: string;
};

type SchemaGraph = {
  attributes: AttributeDeclaration[];
  children: ChildDeclaration[];
  inheritance: Array<{ base: string; derived: string }>;
  namespaces: Array<{ uri: string }>;
  symbols: SchemaSymbol[];
};

const MEASURE_PATTERN = /mm\|cm\|in\|pt\|pc\|pi/u;
const NUMERIC_BUILTIN = /int|long|short|byte|decimal|integer/iu;

type Spelling = "measure" | "numeric" | "percent" | "other";

const buildIndex = (graph: SchemaGraph) => {
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

type Index = ReturnType<typeof buildIndex>;

/**
 * The spellings a type accepts, following union members and restriction bases.
 *
 * A dual-spelling type reports both `numeric` and one of `measure`/`percent`;
 * a type that only ever carries the unit-suffixed string (`ST_TextBulletSizePercent`)
 * reports no `numeric` member and is left alone, because Transitional spells it
 * the same way Strict does.
 */
const spellingsOf = (
  index: Index,
  qualifiedName: string | undefined,
  seen = new Set<string>(),
): Set<Spelling> => {
  const spellings = new Set<Spelling>();
  if (qualifiedName === undefined || seen.has(qualifiedName)) {
    return spellings;
  }
  seen.add(qualifiedName);

  if (qualifiedName.startsWith("{http://www.w3.org/2001/XMLSchema}")) {
    const local = qualifiedName.slice(qualifiedName.indexOf("}") + 1);
    spellings.add(NUMERIC_BUILTIN.test(local) ? "numeric" : "other");
    return spellings;
  }

  const symbol = index.byId.get(`simpleType:${qualifiedName}`);
  if (symbol === undefined) {
    return spellings;
  }
  if (symbol.memberTypes) {
    for (const member of symbol.memberTypes) {
      for (const spelling of spellingsOf(index, member, seen)) {
        spellings.add(spelling);
      }
    }
    return spellings;
  }
  const patterns = (symbol.facets ?? []).filter((facet) => facet.kind === "pattern");
  if (patterns.some((facet) => MEASURE_PATTERN.test(facet.value))) {
    spellings.add("measure");
    return spellings;
  }
  if (patterns.some((facet) => facet.value.includes("%"))) {
    spellings.add("percent");
    return spellings;
  }
  return spellingsOf(index, symbol.base, seen);
};

/** Attributes a complex type declares, following attribute groups and its base. */
const attributesOf = (
  index: Index,
  ownerId: string,
  seen = new Set<string>(),
): Array<{ name: string; type: string }> => {
  if (seen.has(ownerId)) {
    return [];
  }
  seen.add(ownerId);

  const resolved: Array<{ name: string; type: string }> = [];
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
      resolved.push({ name: attribute.name, type: attribute.type });
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
const reachableWmlElementTypes = (
  graph: SchemaGraph,
  index: Index,
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

  for (const root of REBUILT_PART_ROOTS) {
    const element = index.byId.get(`element:{${WML_NAMESPACE}}${root}`);
    reach(root, element?.type);
    if (element?.type) {
      visit(`complexType:${element.type}`);
    }
  }
  return reached;
};

/** Every complex type an element of a given qualified name can be declared with. */
const elementTypes = (graph: SchemaGraph, index: Index): Map<string, Set<string>> => {
  const reachable = reachableWmlElementTypes(graph, index);
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

type Slot = {
  attribute: string | null;
  element: string;
  encoding: { measure?: string; percent?: string };
  namespace: string;
  union: string;
};

const localName = (qualifiedName: string): string =>
  qualifiedName.slice(qualifiedName.indexOf("}") + 1);

const encodingFor = (
  index: Index,
  type: string,
): { measure?: string; percent?: string } | null | undefined => {
  const spellings = spellingsOf(index, type);
  if (!spellings.has("numeric") || (!spellings.has("measure") && !spellings.has("percent"))) {
    return undefined;
  }
  const union = localName(type);
  if (!(union in UNION_ENCODINGS)) {
    throw new GenerateStrictValueEncodingsError({
      message: `${union} spells its value two ways but UNION_ENCODINGS does not say which unit its number is in. Add it, or exclude it with a reason.`,
    });
  }
  return UNION_ENCODINGS[union];
};

const collectSlots = (graph: SchemaGraph, index: Index): Slot[] => {
  const slots: Slot[] = [];
  for (const [key, types] of elementTypes(graph, index)) {
    const separator = key.lastIndexOf(" ");
    const namespace = key.slice(0, separator);
    const element = key.slice(separator + 1);
    for (const type of types) {
      const simple = index.byId.get(`simpleType:${type}`);
      const complex = index.byId.get(`complexType:${type}`);
      const textType = simple ? type : (index.baseOf.get(complex?.id ?? "") ?? undefined);
      if (textType !== undefined && index.byId.get(`simpleType:${textType}`)) {
        const encoding = encodingFor(index, textType);
        if (encoding) {
          slots.push({
            attribute: null,
            element,
            encoding,
            namespace,
            union: localName(textType),
          });
        }
      }
      if (!complex) {
        continue;
      }
      for (const attribute of attributesOf(index, complex.id)) {
        const encoding = encodingFor(index, attribute.type);
        if (encoding) {
          slots.push({
            attribute: attribute.name,
            element,
            encoding,
            namespace,
            union: localName(attribute.type),
          });
        }
      }
    }
  }
  return slots;
};

/**
 * The Strict URI for a Transitional one.
 *
 * ECMA-376 Part 4 republished each markup namespace under `purl.oclc.org/ooxml`
 * with the `/2006` version segment dropped, so the pair is mechanical. The OPC
 * namespaces are not republished: a Strict package addresses its parts and
 * content types exactly as a Transitional one does.
 */
const strictUriFor = (transitional: string): string | undefined => {
  const match = /^http:\/\/schemas\.openxmlformats\.org\/(.+)\/2006\/(.+)$/u.exec(transitional);
  if (match === null || match[1] === "package") {
    return undefined;
  }
  return `http://purl.oclc.org/ooxml/${match[1]}/${match[2]}`;
};

const renderModule = (slots: Slot[], namespacePairs: Array<[string, string]>): string => {
  const bySlot = new Map<string, Slot>();
  for (const slot of slots) {
    const key = `${slot.namespace} ${slot.element}${slot.attribute === null ? "" : ` @${slot.attribute}`}`;
    const existing = bySlot.get(key);
    if (existing && existing.union !== slot.union) {
      throw new GenerateStrictValueEncodingsError({
        message: `${key} resolves to both ${existing.union} and ${slot.union}; the slot key cannot say which unit applies.`,
      });
    }
    bySlot.set(key, slot);
  }

  const entries = [...bySlot.entries()].sort(([left], [right]) => (left < right ? -1 : 1));
  const encodings = entries.map(
    ([key, slot]) => `${key}\t${slot.encoding.measure ?? ""}\t${slot.encoding.percent ?? ""}`,
  );

  const namespaces = namespacePairs
    .sort(([left], [right]) => (left < right ? -1 : 1))
    .map(([strict, transitional]) => `  ["${strict}", "${transitional}"],`);

  return `/**
 * GENERATED FILE — do not edit.
 *
 * How a Strict-spelled measure or percentage is written under a Transitional
 * root, derived from the Transitional schema graph by
 * \`scripts/generate-strict-value-encodings.ts\`. Regenerate with:
 *
 *   bun run generate:strict-value-encodings
 */

import type { MeasureUnit } from "./universalMeasure";

/** The fraction of a percent a Transitional attribute's number counts. */
export type PercentUnit = "fiftiethPercent" | "thousandthPercent" | "wholePercent";

/** The number a slot's dual-spelling type accepts in place of a Strict string. */
export type SlotEncoding = {
  readonly measure?: MeasureUnit;
  readonly percent?: PercentUnit;
};

const MEASURE_UNITS: Readonly<Record<string, MeasureUnit>> = {
  emu: "emu",
  halfPoints: "halfPoints",
  hundredthPoints: "hundredthPoints",
  twips: "twips",
};

const PERCENT_UNITS: Readonly<Record<string, PercentUnit>> = {
  fiftiethPercent: "fiftiethPercent",
  thousandthPercent: "thousandthPercent",
  wholePercent: "wholePercent",
};

/**
 * One slot per line: the slot, a tab, its measure unit, a tab, its percent unit.
 *
 * Text rather than object literals because every package that depends on
 * \`@stll/folio-core\` pays this file's inference cost, and a few hundred
 * literals breach the repository's compiler-workload budget on their own.
 * The slot is \`"<namespace URI> <element local name>"\` for element text and
 * \`"<namespace URI> <element local name> @<attribute local name>"\` for an
 * attribute; an empty column means the type has no spelling of that kind.
 */
const SLOT_TABLE = \`${encodings.join("\n")}\`;

const readSlotTable = (): ReadonlyMap<string, SlotEncoding> => {
  const slots = new Map<string, SlotEncoding>();
  for (const line of SLOT_TABLE.split("\\n")) {
    const [slot, measure, percent] = line.split("\\t");
    if (slot === undefined) {
      continue;
    }
    const measureUnit = measure === undefined ? undefined : MEASURE_UNITS[measure];
    const percentUnit = percent === undefined ? undefined : PERCENT_UNITS[percent];
    slots.set(slot, {
      ...(measureUnit === undefined ? {} : { measure: measureUnit }),
      ...(percentUnit === undefined ? {} : { percent: percentUnit }),
    });
  }
  return slots;
};

/** Slots whose Transitional type spells one value two ways. */
export const TRANSITIONAL_SLOT_ENCODINGS: ReadonlyMap<string, SlotEncoding> = readSlotTable();

const NAMESPACE_PAIRS: readonly (readonly [strict: string, transitional: string])[] = [
${namespaces.join("\n")}
];

/** Every Strict namespace URI a WordprocessingML part can carry, and its Transitional pair. */
export const TRANSITIONAL_NAMESPACE_BY_STRICT_URI: ReadonlyMap<string, string> = new Map(
  NAMESPACE_PAIRS,
);
`;
};

const digest = (value: string): string => createHash("sha256").update(value).digest("hex");

const main = async (): Promise<void> => {
  const mode = process.argv.at(2) ?? "write";
  if (mode !== "check" && mode !== "write") {
    throw new GenerateStrictValueEncodingsError({
      message: "Usage: bun scripts/generate-strict-value-encodings.ts [check|write]",
    });
  }

  const graph = JSON.parse(await readFile(GRAPH_PATH, "utf8")) as SchemaGraph;
  const index = buildIndex(graph);
  const slots = collectSlots(graph, index);

  const namespacePairs: Array<[string, string]> = [];
  for (const { uri } of graph.namespaces) {
    const strict = strictUriFor(uri);
    if (strict !== undefined) {
      namespacePairs.push([strict, uri]);
    }
  }

  if (!slots.some((slot) => slot.element === "tcW" && slot.attribute === "w")) {
    throw new GenerateStrictValueEncodingsError({
      message: "w:tcW/@w:w must be a twips slot; the schema graph or the derivation drifted.",
    });
  }

  const rendered = renderModule(slots, namespacePairs);
  if (mode === "write") {
    await writeFile(OUTPUT_PATH, rendered, "utf8");
    process.stdout.write(
      `strictValueEncodings.gen.ts written (${new Set(slots.map((s) => `${s.namespace} ${s.element} ${s.attribute}`)).size} slots, ${namespacePairs.length} namespaces)\n`,
    );
    return;
  }

  const existing = await readFile(OUTPUT_PATH, "utf8").catch(() => null);
  if (existing === null || digest(existing) !== digest(rendered)) {
    throw new GenerateStrictValueEncodingsError({
      message:
        "strictValueEncodings.gen.ts is stale. Run `bun run generate:strict-value-encodings`.",
    });
  }
  process.stdout.write("strictValueEncodings.gen.ts is up to date\n");
};

await main();
