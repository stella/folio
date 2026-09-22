/**
 * Representative values for a schema simple type.
 *
 * The survival law has to write something into every attribute it tests, and
 * what it writes decides what the law proves. A single value per attribute
 * proves only that one path; the census therefore draws a small, deterministic
 * set from the type itself — every enumeration member, both spellings of a
 * dual-spelled measure, every `ST_OnOff` token, `auto` as well as a real hex
 * colour, and the bounds an integer facet declares. The set is derived, so a
 * schema refresh that adds an enumeration member adds a case here without
 * anybody noticing it has to.
 *
 * The first value of the set is the *representative*: the one the exhaustive
 * pair sweep uses, so its cost stays linear in the number of pairs. The rest
 * are what the value sweep and the property test draw from.
 *
 * A type may have more than one representative. `ST_OnOff` has two meanings,
 * not six values: an on, and an off that overrides an inherited setting. One
 * representative would measure whichever of them it happened to be and say
 * nothing about the other, so the meanings the reserved-value registry already
 * names are drawn out as representatives of their own.
 */

import { RESERVED_VALUE_NAMESPACE_URIS } from "../../../specifications/reserved-values/disposition";
import { reservedValueEntries } from "../../../specifications/reserved-values/registry";
import type { OoxmlSchemaGraph } from "../../generate-ooxml-schema-graph";
import type { AttributeSlot, SchemaIndex } from "./schemaSpace";

type SchemaSymbol = OoxmlSchemaGraph["symbols"][number];

const XSD = "http://www.w3.org/2001/XMLSchema";

/** Every `xs:boolean`-and-token spelling `ST_OnOff` accepts. */
const ON_OFF_VALUES = ["true", "false", "1", "0", "on", "off"] as const;

/** A measure type that accepts both a bare number and a unit-suffixed string. */
const MEASURE_PATTERN = /mm\|cm\|in\|pt\|pc\|pi/u;

const NUMERIC_BUILTINS = new Set([
  "byte",
  "decimal",
  "int",
  "integer",
  "long",
  "negativeInteger",
  "nonNegativeInteger",
  "nonPositiveInteger",
  "positiveInteger",
  "short",
  "unsignedByte",
  "unsignedInt",
  "unsignedLong",
  "unsignedShort",
]);

const builtinLocalName = (qualifiedName: string): string | undefined =>
  qualifiedName.startsWith(`{${XSD}}`)
    ? qualifiedName.slice(qualifiedName.indexOf("}") + 1)
    : undefined;

type Facets = ReadonlyArray<{ kind: string; value: string }>;

const facetValue = (facets: Facets, kind: string): string | undefined =>
  facets.find((facet) => facet.kind === kind)?.value;

/**
 * Integer values that exercise the bounds the schema declares.
 *
 * `1` leads because it is the one integer no falsy test swallows: a
 * representative of `0` would let a serializer that writes `if (value)` read as
 * dropping every slot it holds. `0` and `-1` follow, because a sign or index
 * confusion shows up on them, then whatever minimum and maximum the facets
 * declare, where a clamp does.
 */
const integerValues = (facets: Facets, builtin: string): string[] => {
  const declaredMin = facetValue(facets, "minInclusive") ?? facetValue(facets, "minExclusive");
  const declaredMax = facetValue(facets, "maxInclusive") ?? facetValue(facets, "maxExclusive");
  const signed =
    !builtin.startsWith("unsigned") &&
    !builtin.startsWith("nonNegative") &&
    builtin !== "positiveInteger";
  const candidates = [
    "1",
    "0",
    ...(signed ? ["-1"] : []),
    ...(declaredMin === undefined ? [] : [declaredMin]),
    ...(declaredMax === undefined ? [] : [declaredMax]),
  ];
  return [...new Set(candidates)].filter(
    (value) =>
      (declaredMin === undefined || BigInt(value) >= BigInt(declaredMin)) &&
      (declaredMax === undefined || BigInt(value) <= BigInt(declaredMax)),
  );
};

/** Hex of the byte length the type's `length` facet declares; `FF0000` when it declares none. */
const hexValues = (facets: Facets): string[] => {
  const bytes = Number.parseInt(facetValue(facets, "length") ?? "3", 10);
  const width = (Number.isFinite(bytes) ? bytes : 3) * 2;
  return [...new Set(["FF0000".padEnd(width, "0").slice(0, width), "0".repeat(width)])];
};

/** A token that survives every normaliser: no whitespace, no escaping, no locale. */
const PLAIN_TOKEN = "folio1";

const DATE_VALUE = "2024-01-01T00:00:00Z";

type ValueSet = {
  /** Deterministic, ordered, deduplicated. The first member is the representative. */
  values: readonly string[];
  /** Why the set looks the way it does; the census reports it for an unrepresentable slot. */
  kind: "enumeration" | "on-off" | "hex" | "measure" | "integer" | "string" | "unknown";
};

const NO_VALUES: ValueSet = { values: [], kind: "unknown" };

let reservedEntries: ReturnType<typeof reservedValueEntries> | undefined;

const allReservedEntries = (): ReturnType<typeof reservedValueEntries> => {
  reservedEntries ??= reservedValueEntries();
  return reservedEntries;
};

const isReservedValuePrefix = (
  prefix: string,
): prefix is keyof typeof RESERVED_VALUE_NAMESPACE_URIS => prefix in RESERVED_VALUE_NAMESPACE_URIS;

const matchesReservedSlot = (slot: AttributeSlot, registered: string): boolean => {
  const colon = registered.indexOf(":");
  const at = registered.indexOf("@");
  if (colon < 1 || at < colon + 2) {
    return false;
  }
  const prefix = registered.slice(0, colon);
  const namespace = isReservedValuePrefix(prefix)
    ? RESERVED_VALUE_NAMESPACE_URIS[prefix]
    : undefined;
  return (
    namespace === slot.container.element.namespace &&
    registered.slice(colon + 1, at) === slot.container.element.name &&
    registered.slice(at + 1) === slot.attribute.name
  );
};

/** Why a registered sentinel cannot be a value subject in a valid model. */
export const unrepresentableReservedValue = (
  slot: AttributeSlot,
  value: string,
): string | undefined => {
  for (const { disposition } of allReservedEntries()) {
    if (
      disposition === "no-reserved-value" ||
      disposition.disposition !== "unrepresentable" ||
      !disposition.sentinel.split("|").includes(value) ||
      !disposition.slot.split("|").some((registered) => matchesReservedSlot(slot, registered))
    ) {
      continue;
    }
    return `reserved value ${value} is excluded by ${disposition.carrier}`;
  }
  return undefined;
};

const resolveSymbol = (index: SchemaIndex, qualifiedName: string): SchemaSymbol | undefined =>
  index.byId.get(`simpleType:${qualifiedName}`);

/**
 * The values a simple type accepts, following unions and restriction bases.
 *
 * `ST_OnOff` is recognised by its union of `xs:boolean` and the on/off token
 * enumeration rather than by name, so a type that spells the same idea under
 * another name is covered too.
 */
export const valuesForType = (
  index: SchemaIndex,
  qualifiedName: string | undefined,
  seen = new Set<string>(),
): ValueSet => {
  if (qualifiedName === undefined || seen.has(qualifiedName)) {
    return NO_VALUES;
  }
  seen.add(qualifiedName);

  const builtin = builtinLocalName(qualifiedName);
  if (builtin !== undefined) {
    if (builtin === "boolean") {
      return { values: ["true", "false", "1", "0"], kind: "on-off" };
    }
    if (NUMERIC_BUILTINS.has(builtin)) {
      return { values: integerValues([], builtin), kind: "integer" };
    }
    if (builtin === "hexBinary") {
      return { values: hexValues([]), kind: "hex" };
    }
    if (builtin === "dateTime") {
      return { values: [DATE_VALUE], kind: "string" };
    }
    return { values: [PLAIN_TOKEN], kind: "string" };
  }

  const symbol = resolveSymbol(index, qualifiedName);
  if (symbol === undefined) {
    return NO_VALUES;
  }

  if (symbol.enumValues && symbol.enumValues.length > 0) {
    return { values: [...new Set(symbol.enumValues)], kind: "enumeration" };
  }

  if (symbol.memberTypes && symbol.memberTypes.length > 0) {
    const members = symbol.memberTypes.map((member) => valuesForType(index, member, seen));
    const merged = [...new Set(members.flatMap(({ values }) => values))];
    const isOnOff = ON_OFF_VALUES.every((value) => merged.includes(value));
    if (isOnOff) {
      return { values: [...ON_OFF_VALUES], kind: "on-off" };
    }
    const memberKinds = new Set(members.map(({ kind: memberKind }) => memberKind));
    const dominant = (["hex", "measure"] as const).find((candidate) => memberKinds.has(candidate));
    const kind = dominant ?? members.find(({ values }) => values.length > 0)?.kind ?? "unknown";
    return merged.length === 0 ? NO_VALUES : { values: merged, kind };
  }

  const facets = symbol.facets ?? [];
  const patterns = facets.filter((facet) => facet.kind === "pattern").map(({ value }) => value);
  if (patterns.some((pattern) => MEASURE_PATTERN.test(pattern))) {
    // Both spellings of the same length: Strict writes the suffixed form and
    // Transitional the bare twip count, and folio has to read either.
    return { values: ["72pt", "1in", "2.54cm"], kind: "measure" };
  }
  if (patterns.some((pattern) => pattern.includes("%"))) {
    return { values: ["50%", "0%", "100%"], kind: "measure" };
  }
  const base = symbol.base;
  if (base === undefined) {
    return { values: [PLAIN_TOKEN], kind: "string" };
  }
  const inherited = valuesForType(index, base, seen);
  if (inherited.kind === "integer") {
    return { values: integerValues(facets, builtinLocalName(base) ?? "integer"), kind: "integer" };
  }
  if (inherited.kind === "hex") {
    return { values: hexValues(facets), kind: "hex" };
  }
  return inherited;
};

/**
 * A spelling of a reserved value that identifies the type declaring it.
 *
 * A sentinel in `specifications/reserved-values` is one reserved meaning and
 * its `|`-separated alternatives are spellings of it: `0|false|off` is the one
 * "not set" a toggle has, `nil|none` the one "no border". Only the alternatives
 * that are tokens are usable here. A numeric one — `0`, `-1`, `240` — names a
 * point in a lexical space every integer type admits, so matching on it would
 * hand `ST_TwipsMeasure` and `ST_Coordinate` a second representative apiece for
 * a decision the registry recorded about `w:numId`. A token is only legal where
 * the type that names it is in play.
 */
const RESERVED_TOKEN = /^[A-Za-z][A-Za-z0-9]*$/u;

let reservedMeanings: string[][] | undefined;

/** One entry per reserved meaning: the token spellings the registry names for it. */
const reservedTokenGroups = (): string[][] => {
  if (reservedMeanings !== undefined) {
    return reservedMeanings;
  }
  const seen = new Set<string>();
  const groups: string[][] = [];
  for (const { disposition } of reservedValueEntries()) {
    if (disposition === "no-reserved-value" || seen.has(disposition.sentinel)) {
      continue;
    }
    seen.add(disposition.sentinel);
    const tokens = disposition.sentinel.split("|").filter((token) => RESERVED_TOKEN.test(token));
    if (tokens.length > 0) {
      groups.push(tokens);
    }
  }
  reservedMeanings = groups;
  return groups;
};

/**
 * The values the exhaustive pair sweep writes, one per meaning the type has.
 *
 * The first is the ordinary one, and it leads because `allSubjects` and the
 * container contract key a pair by its slot alone: the pair sweep's cost stays
 * linear in the number of pairs. `auto` leads `ST_HexColor`'s member list and
 * is a poor value to lead with — a slot that only ever survived as `auto` would
 * read as surviving — so the hex is preferred there; every other type's first
 * member is already its safest.
 *
 * After it come the reserved meanings the type spells, one value each. They are
 * the values that mean the opposite of an ordinary one, and the census measures
 * every one of them whatever the others did.
 */
export const representativeValues = (
  index: SchemaIndex,
  qualifiedName: string | undefined,
): readonly string[] => {
  const { values, kind } = valuesForType(index, qualifiedName);
  const ordinary =
    kind === "hex"
      ? (values.find((value) => /^[0-9A-Fa-f]+$/u.test(value)) ?? values.at(0))
      : values.at(0);
  if (ordinary === undefined) {
    return [];
  }
  const representatives = [ordinary];
  for (const group of reservedTokenGroups()) {
    const spelling = group.find((token) => values.includes(token));
    if (spelling !== undefined && !representatives.includes(spelling)) {
      representatives.push(spelling);
    }
  }
  return representatives;
};

/** The one value a pair is keyed on: the first of {@link representativeValues}. */
export const representativeValue = (
  index: SchemaIndex,
  qualifiedName: string | undefined,
): string | undefined => representativeValues(index, qualifiedName).at(0);
