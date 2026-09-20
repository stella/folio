#!/usr/bin/env bun

/**
 * Every OOXML slot that can carry a reserved value is either registered or excluded.
 *
 * `specifications/reserved-values` is total over the *model*: a field
 * added without a decision fails typecheck. This check covers the other
 * direction, over the *schema*: a slot the format declares but the model does
 * not reach is invisible to the compiler, so it is derived mechanically from
 * the committed schema graph and must be named by a registry entry or by an
 * explicit exclusion carrying a reason.
 *
 * A candidate slot is one of three kinds:
 *   enum     an attribute whose simple type enumerates a reserved token
 *            (auto, nil, none, clear, nothing, continue, baseline, default, custom)
 *   default  an attribute the schema gives a default value, so an absent
 *            attribute and one written with its default mean the same thing
 *   prose    a numeric sentinel the schema cannot express, curated below
 *
 * A second direction, over the *enumerations*: naming a slot is not the same as
 * reading it. A registry entry can describe a slot whose model enumeration has
 * drifted from the schema's, and then the decision covers tokens that never
 * reach it — `ST_ThemeColor` was such an entry, wrong about six of its sixteen
 * members. So every enumerating type a registered slot can carry must also be
 * bound to what the model accepts, and the two compared member for member.
 *
 * No network: the graph is committed, the way `generate:strict-value-encodings`
 * reads it.
 *
 * Usage: bun scripts/check-reserved-value-coverage.ts
 */

import { TaggedError } from "better-result";

import { RESERVED_VALUE_NAMESPACE_URIS } from "../specifications/reserved-values/disposition";
import { reservedValueSlotKeys } from "../specifications/reserved-values/registry";
import { keyDifferences } from "./lib/exact-object";
import {
  attributesOf,
  buildIndex,
  elementTypes,
  type Index,
  INLINE_NAMESPACES,
  type SchemaGraph,
  loadSchemaGraph,
  localName,
  REBUILT_PART_ROOTS,
  slotKey,
} from "./lib/ooxml-schema-graph";
import {
  expandSimpleType,
  narrowedEnumOptions,
  narrowedEnumSiteByType,
  RESERVED_VALUE_SLOT_ENUMERATIONS,
  type SlotEnumerationDecision,
} from "./lib/narrowed-enum-schema-types";
import {
  RESERVED_VALUE_EXCLUSIONS,
  type ReservedValueExclusionGroup,
} from "./lib/reserved-value-exclusions";

class ReservedValueCoverageError extends TaggedError("ReservedValueCoverageError")<{
  message: string;
}> {}

/**
 * Tokens whose presence in an enumeration marks the slot as carrying a reserved
 * value: each of them names "no value", "the value is elsewhere", or "the same
 * as the level above", never a value of its own.
 */
const RESERVED_ENUM_TOKENS = new Set([
  "auto",
  "baseline",
  "clear",
  "continue",
  "custom",
  "default",
  "nil",
  "none",
  "nothing",
  // `ST_OnOff` accepts `off` as an explicit "not set", which is not the same as
  // the element being absent, and two levels of the style hierarchy combine by
  // XOR rather than by override.
  "off",
]);

/**
 * The parts folio reads into the model.
 *
 * `REBUILT_PART_ROOTS` covers what folio writes back; a reserved value has to
 * be decided wherever folio *reads* one, which also takes in the numbering and
 * style definitions it resolves against.
 */
const RESERVED_VALUE_PART_ROOTS: readonly string[] = [...REBUILT_PART_ROOTS, "numbering", "styles"];

/**
 * Numeric sentinels the schema cannot express: `w:numId`, `w:ilvl`,
 * `w:outlineLvl`, `w:gridSpan`, `w:start` and `w:lvlRestart` are all
 * `CT_DecimalNumber` over an unfacetted `xs:integer`, and the note ids are
 * plain integers too. Each is pinned by a record in `specifications/evidence`.
 */
const PROSE_SLOTS: readonly string[] = [
  "w:numId@val",
  "w:ilvl@val",
  "w:outlineLvl@val",
  "w:gridSpan@val",
  "w:start@val",
  "w:lvlRestart@val",
  "w:startOverride@val",
  "w:lvl@ilvl",
  "w:lvlOverride@ilvl",
  "w:num@numId",
  "w:footnote@id",
  "w:endnote@id",
  "w:footnoteReference@id",
  "w:endnoteReference@id",
  "w:spacing@line",
  "w:ind@firstLine",
  "w:ind@hanging",
  // Absent means `nextPage`; the schema records no default.
  "w:type@val",
  // The number is meaningless under `@w:type="auto"`, which no schema can say.
  "w:tblW@w",
  "w:tcW@w",
  "w:tblCellSpacing@w",
  "w:tblInd@w",
  "w:wBefore@w",
  "w:wAfter@w",
  // `autofit` and `never` are named members, but the behaviour they name is prose.
  "w:tblLayout@type",
  "w:tblOverlap@val",
  // A style reference nothing resolves falls back to the type's default style.
  "w:pStyle@val",
  "w:rStyle@val",
  "w:tblStyle@val",
  "w:numStyleLink@val",
  "w:styleLink@val",
];

type Candidate = { slot: string; kind: "enum" | "default" | "prose"; detail: string };

/** Every enum token a type accepts, following union members and restriction bases. */
const enumTokensOf = (
  index: Index,
  type: string | undefined,
  seen = new Set<string>(),
): string[] => {
  if (type === undefined || seen.has(type)) {
    return [];
  }
  seen.add(type);
  const symbol = index.byId.get(`simpleType:${type}`);
  if (symbol === undefined) {
    return [];
  }
  if (symbol.enumValues && symbol.enumValues.length > 0) {
    return symbol.enumValues;
  }
  const tokens = (symbol.memberTypes ?? []).flatMap((member) => enumTokensOf(index, member, seen));
  return tokens.length > 0 ? tokens : enumTokensOf(index, symbol.base, seen);
};

const collectCandidates = (graph: SchemaGraph, index: Index): Candidate[] => {
  const candidates = new Map<string, Candidate>();
  const add = (candidate: Candidate): void => {
    if (!candidates.has(candidate.slot)) {
      candidates.set(candidate.slot, candidate);
    }
  };

  for (const [key, types] of elementTypes(graph, index, RESERVED_VALUE_PART_ROOTS)) {
    const separator = key.lastIndexOf(" ");
    const namespace = key.slice(0, separator);
    const element = key.slice(separator + 1);
    for (const type of types) {
      const complex = index.byId.get(`complexType:${type}`);
      if (!complex) {
        continue;
      }
      for (const attribute of attributesOf(index, complex.id)) {
        const slot = slotKey(namespace, element, attribute.name);
        const reserved = enumTokensOf(index, attribute.type).filter((token) =>
          RESERVED_ENUM_TOKENS.has(token),
        );
        if (reserved.length > 0) {
          add({ slot, kind: "enum", detail: reserved.toSorted().join(", ") });
          continue;
        }
        if (attribute.default !== undefined) {
          add({ slot, kind: "default", detail: `default "${attribute.default}"` });
        }
      }
    }
  }
  return [...candidates.values()].toSorted((left, right) => left.slot.localeCompare(right.slot));
};

/** `"w:tcW@type"` -> `"<wml uri> tcW @type"`. */
const expandSlot = (slot: string): string => {
  const colon = slot.indexOf(":");
  const prefix = slot.slice(0, colon);
  const uri = (RESERVED_VALUE_NAMESPACE_URIS as Record<string, string>)[prefix];
  if (uri === undefined) {
    throw new ReservedValueCoverageError({
      message: `Slot "${slot}" uses the namespace prefix "${prefix}", which RESERVED_VALUE_NAMESPACE_URIS does not declare.`,
    });
  }
  const rest = slot.slice(colon + 1);
  const at = rest.indexOf("@");
  return at < 0 ? slotKey(uri, rest, null) : slotKey(uri, rest.slice(0, at), rest.slice(at + 1));
};

/**
 * Every attribute slot the schema declares in a namespace a WordprocessingML
 * part can carry, ignoring reachability.
 *
 * A registry or exclusion slot outside this set is a typo the compiler cannot
 * see: the string names nothing in the format. Elements are recorded without an
 * attribute too, because a few reserved values are an element's presence
 * (`a:noFill`, `a:noAutofit`) rather than an attribute's value.
 */
const declaredSlots = (graph: SchemaGraph, index: Index): Set<string> => {
  const declared = new Set<string>();
  for (const symbol of graph.symbols) {
    if (symbol.kind !== "element" || symbol.type === undefined) {
      continue;
    }
    if (!INLINE_NAMESPACES.has(symbol.namespace)) {
      continue;
    }
    declared.add(slotKey(symbol.namespace, symbol.name, null));
    const complex = index.byId.get(`complexType:${symbol.type}`);
    if (complex) {
      for (const attribute of attributesOf(index, complex.id)) {
        declared.add(slotKey(symbol.namespace, symbol.name, attribute.name));
      }
    }
  }
  for (const child of graph.children) {
    if (child.kind !== "element") {
      continue;
    }
    const element = child.ref ? index.byId.get(`element:${child.ref}`) : undefined;
    const namespace = element?.namespace ?? child.namespace;
    const name = element?.name ?? child.name;
    const type = element?.type ?? child.type;
    if (name === undefined || namespace === undefined || !INLINE_NAMESPACES.has(namespace)) {
      continue;
    }
    declared.add(slotKey(namespace, name, null));
    const complex = type === undefined ? undefined : index.byId.get(`complexType:${type}`);
    if (complex) {
      for (const attribute of attributesOf(index, complex.id)) {
        declared.add(slotKey(namespace, name, attribute.name));
      }
    }
  }
  return declared;
};

const excludedSlots = (groups: readonly ReservedValueExclusionGroup[]): Map<string, string> => {
  const excluded = new Map<string, string>();
  for (const group of groups) {
    for (const slot of group.slots) {
      const expanded = expandSlot(slot);
      if (excluded.has(expanded)) {
        throw new ReservedValueCoverageError({
          message: `Slot "${slot}" is excluded twice; one entry is dead.`,
        });
      }
      excluded.set(expanded, group.reason);
    }
  }
  return excluded;
};

/**
 * Every enumerating simple type a registered slot can carry, and the slots that
 * carry it.
 *
 * `enumTokensOf` follows union members and restriction bases, so a type that
 * only enumerates through its base is included the way the candidate walk
 * includes it.
 */
const enumeratingTypesOfRegisteredSlots = (
  graph: SchemaGraph,
  index: Index,
  registered: ReadonlySet<string>,
): Map<string, { slots: Set<string>; tokens: readonly string[] }> => {
  const reached = new Map<string, { slots: Set<string>; tokens: readonly string[] }>();
  for (const [key, types] of elementTypes(graph, index, RESERVED_VALUE_PART_ROOTS)) {
    const separator = key.lastIndexOf(" ");
    const namespace = key.slice(0, separator);
    const element = key.slice(separator + 1);
    for (const type of types) {
      const complex = index.byId.get(`complexType:${type}`);
      if (!complex) {
        continue;
      }
      for (const attribute of attributesOf(index, complex.id)) {
        const slot = slotKey(namespace, element, attribute.name);
        if (!registered.has(slot)) {
          continue;
        }
        const tokens = enumTokensOf(index, attribute.type);
        if (tokens.length === 0) {
          continue;
        }
        const entry = reached.get(attribute.type) ?? { slots: new Set<string>(), tokens };
        entry.slots.add(slot);
        reached.set(attribute.type, entry);
      }
    }
  }
  return reached;
};

const difference = (left: readonly string[], right: readonly string[]): string[] =>
  left.filter((value) => !right.includes(value)).toSorted();

/**
 * The model's accepted token set against the schema's, for every enumerating
 * type a registered slot reaches.
 *
 * Three verdicts, and only the first two pass: the model's set is the
 * enumeration; the divergence is the one the binding records; or the type is
 * bound to nothing, which is the failure this check exists for.
 */
const enumerationMismatches = (
  reached: ReadonlyMap<string, { slots: Set<string>; tokens: readonly string[] }>,
): string[] => {
  const problems: string[] = [];
  const bySite = narrowedEnumSiteByType();
  const decisions = new Map<string, SlotEnumerationDecision>(
    Object.entries(RESERVED_VALUE_SLOT_ENUMERATIONS).map(([qualified, decision]) => [
      expandSimpleType(qualified),
      decision,
    ]),
  );

  for (const [type, { slots, tokens }] of [...reached.entries()].toSorted()) {
    const named = `${localName(type)} (${[...slots].toSorted().join(", ")})`;

    const picklist = bySite.get(type);
    if (picklist) {
      const options = narrowedEnumOptions(picklist.site);
      const missing = difference(tokens, options);
      const extra = difference(options, tokens);
      const { binding } = picklist;
      if (binding.kind === "matches" && (missing.length > 0 || extra.length > 0)) {
        problems.push(
          `  ${named}: ${picklist.site} is recorded as matching ${binding.simpleType}, ` +
            `but omits ${missing.join(", ") || "nothing"} and adds ${extra.join(", ") || "nothing"}.`,
        );
        continue;
      }
      if (binding.kind === "keeps-raw" && extra.length > 0) {
        problems.push(
          `  ${named}: ${picklist.site} accepts ${extra.join(", ")}, which ${binding.simpleType} does not declare.`,
        );
        continue;
      }
      if (
        binding.kind === "diverges" &&
        (missing.join() !== [...binding.missing].toSorted().join() ||
          extra.join() !== [...binding.extra].toSorted().join())
      ) {
        problems.push(
          `  ${named}: ${picklist.site} diverges from ${binding.simpleType} by ` +
            `missing ${missing.join(", ") || "nothing"} / extra ${extra.join(", ") || "nothing"}, ` +
            "which is not what the binding records.",
        );
      }
      continue;
    }

    const decision = decisions.get(type);
    if (decision === undefined) {
      problems.push(
        `  ${named}: no model enumeration is bound to it. Record one in ` +
          "scripts/lib/narrowed-enum-schema-types.ts, or say there why the model does not " +
          "hold its tokens.",
      );
      continue;
    }
    if (decision.kind === "excluded") {
      continue;
    }
    const missing = difference(tokens, decision.tokens);
    const extra = difference(decision.tokens, tokens);
    if (missing.length > 0 || extra.length > 0) {
      problems.push(
        `  ${named}: ${decision.source} omits ${missing.join(", ") || "nothing"} ` +
          `and adds ${extra.join(", ") || "nothing"}.`,
      );
    }
  }
  return problems;
};

const main = async (): Promise<void> => {
  const graph = await loadSchemaGraph();
  const index = buildIndex(graph);

  const candidates = collectCandidates(graph, index);
  const candidateSlots = new Set(candidates.map(({ slot }) => slot));
  for (const slot of PROSE_SLOTS) {
    const expanded = expandSlot(slot);
    if (!candidateSlots.has(expanded)) {
      candidates.push({ slot: expanded, kind: "prose", detail: "numeric sentinel" });
      candidateSlots.add(expanded);
    }
  }

  const registered = new Set(reservedValueSlotKeys().map(expandSlot));
  const excluded = excludedSlots(RESERVED_VALUE_EXCLUSIONS);

  // A registry or exclusion slot the schema does not declare is a typo the
  // compiler cannot see; an exclusion for a slot that is no longer a candidate
  // is dead weight.
  const declared = declaredSlots(graph, index);
  const undeclared = [...registered].filter((slot) => !declared.has(slot)).toSorted();

  const covered = [...new Set([...registered, ...excluded.keys()])].toSorted();
  const { missing, extra } = keyDifferences(covered, [...candidateSlots].toSorted());
  const detailOf = new Map(candidates.map((candidate) => [candidate.slot, candidate]));
  const uncovered = missing
    .map((slot) => detailOf.get(slot))
    .filter((candidate): candidate is Candidate => candidate !== undefined);
  const strayExcluded = extra.filter((slot) => excluded.has(slot));

  const reachedEnumerations = enumeratingTypesOfRegisteredSlots(graph, index, registered);
  const mismatched = enumerationMismatches(reachedEnumerations);

  const byKind = (kind: Candidate["kind"]): number =>
    candidates.filter((candidate) => candidate.kind === kind).length;

  process.stdout.write(
    `reserved-value coverage: ${candidates.length} candidate slots ` +
      `(${byKind("enum")} enum, ${byKind("default")} XSD default, ${byKind("prose")} prose); ` +
      `${registered.size} registered, ${excluded.size} excluded across ${RESERVED_VALUE_EXCLUSIONS.length} groups; ` +
      `${reachedEnumerations.size} enumerations bound to the model\n`,
  );

  if (
    uncovered.length === 0 &&
    undeclared.length === 0 &&
    strayExcluded.length === 0 &&
    mismatched.length === 0
  ) {
    return;
  }

  const lines: string[] = [];
  if (uncovered.length > 0) {
    lines.push(
      `${uncovered.length} slot(s) carry a reserved value that no registry entry names and no exclusion covers:`,
      ...uncovered.map(({ slot, kind, detail }) => `  [${kind}] ${slot} (${detail})`),
      "",
      "Record the decision on the model field that carries the slot, or add the slot to a group",
      "in scripts/lib/reserved-value-exclusions.ts with a reason folio can be held to.",
    );
  }
  if (undeclared.length > 0) {
    lines.push(
      "",
      `${undeclared.length} registry slot(s) name no attribute the schema declares:`,
      ...undeclared.map((slot) => `  ${slot}`),
    );
  }
  if (strayExcluded.length > 0) {
    lines.push(
      "",
      `${strayExcluded.length} exclusion(s) are dead; delete them:`,
      ...strayExcluded.map((slot) => `  ${slot} (${excluded.get(slot) ?? ""})`),
    );
  }
  if (mismatched.length > 0) {
    lines.push(
      "",
      `${mismatched.length} registered slot enumeration(s) do not match the schema:`,
      ...mismatched,
    );
  }
  throw new ReservedValueCoverageError({ message: lines.join("\n") });
};

export {
  collectCandidates,
  enumeratingTypesOfRegisteredSlots,
  enumerationMismatches,
  enumTokensOf,
  excludedSlots,
  expandSlot,
};

if (import.meta.main) {
  await main();
}
