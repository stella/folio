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

import {
  attributesOf,
  buildIndex,
  elementTypes,
  type Index,
  localName,
  loadSchemaGraph,
  type SchemaGraph,
} from "./lib/ooxml-schema-graph";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/core/src/docx/strictValueEncodings.gen.ts");

class GenerateStrictValueEncodingsError extends TaggedError("GenerateStrictValueEncodingsError")<{
  message: string;
}> {}

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

const MEASURE_PATTERN = /mm\|cm\|in\|pt\|pc\|pi/u;
const NUMERIC_BUILTIN = /int|long|short|byte|decimal|integer/iu;

type Spelling = "measure" | "numeric" | "percent" | "other";

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

type Slot = {
  attribute: string | null;
  element: string;
  encoding: { measure?: string; percent?: string };
  namespace: string;
  union: string;
};

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
  const encodings = entries.map(([key, slot]) => {
    const columns = [key, slot.encoding.measure ?? "", slot.encoding.percent ?? ""];
    while (columns.at(-1) === "") {
      columns.pop();
    }
    return columns.join("\t");
  });

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

  const graph = await loadSchemaGraph();
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
