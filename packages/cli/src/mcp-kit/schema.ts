/**
 * Advertised schemas. Whatever `tools/list` returns is resent on every turn,
 * so a listed schema carries the shape and nothing else: no descriptions, no
 * titles, no examples. Guidance lives in `describe_capability`.
 */

import type { JsonSchema, ListedTool } from "./types";

const ANNOTATION_KEYWORDS: ReadonlySet<string> = new Set([
  "$comment",
  "$schema",
  "default",
  "description",
  "examples",
  "title",
]);

/** Keywords whose value is a map of property name to schema. */
const SCHEMA_MAP_KEYWORDS: ReadonlySet<string> = new Set([
  "$defs",
  "definitions",
  "patternProperties",
  "properties",
]);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/**
 * `describedDepth` keeps `description` on schemas that many property maps
 * down (1: the tool's own arguments) and strips it below, where a large
 * nested schema repeats itself.
 */
const compactNode = (node: unknown, depth: number, describedDepth: number): unknown => {
  if (Array.isArray(node)) return node.map((item) => compactNode(item, depth, describedDepth));
  if (!isRecord(node)) return node;
  const compacted: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(node)) {
    const keptDescription = key === "description" && depth > 0 && depth <= describedDepth;
    if (ANNOTATION_KEYWORDS.has(key) && !keptDescription) continue;
    // A bound at the largest safe integer says only "an integer".
    if (key === "maximum" && value === Number.MAX_SAFE_INTEGER) continue;
    compacted[key] =
      SCHEMA_MAP_KEYWORDS.has(key) && isRecord(value)
        ? Object.fromEntries(
            Object.entries(value).map(([name, schema]) => [
              name,
              compactNode(schema, depth + 1, describedDepth),
            ]),
          )
        : compactNode(value, depth, describedDepth);
  }
  return compacted;
};

export type CompactSchemaOptions = {
  /** Keep descriptions this many property levels deep; 0 (the default) keeps none. */
  describedDepth?: number;
};

/** The schema without annotation keywords, descriptions kept only as deep as asked. */
export const compactSchema = (
  schema: JsonSchema,
  { describedDepth = 0 }: CompactSchemaOptions = {},
): JsonSchema => {
  const compacted = compactNode(schema, 0, describedDepth);
  return isRecord(compacted) ? compacted : {};
};

type Occurrence = { count: number; name: string; bytes: number };

/** Count every object subschema below the root by its serialization, naming each by where it first appears. */
const countSubschemas = (
  node: unknown,
  name: string,
  seen: Map<string, Occurrence>,
  isRoot: boolean,
): void => {
  if (Array.isArray(node)) {
    for (const item of node) countSubschemas(item, name, seen, false);
    return;
  }
  if (!isRecord(node)) return;
  if (!isRoot) {
    const key = JSON.stringify(node);
    const occurrence = seen.get(key);
    if (occurrence === undefined) seen.set(key, { count: 1, name, bytes: key.length });
    else occurrence.count += 1;
  }
  for (const [key, value] of Object.entries(node)) {
    if (SCHEMA_MAP_KEYWORDS.has(key) && isRecord(value)) {
      for (const [child, schema] of Object.entries(value))
        countSubschemas(schema, child, seen, false);
    } else {
      countSubschemas(value, name, seen, false);
    }
  }
};

/**
 * Hoist every object subschema that occurs more than once and serializes to
 * at least `minBytes` into `$defs`, referenced by `$ref`: a schema that
 * repeats one shape for several properties states it once.
 */
export const hoistRepeatedSchemas = (
  schema: JsonSchema,
  { minBytes = 200 }: { minBytes?: number } = {},
): JsonSchema => {
  const seen = new Map<string, Occurrence>();
  countSubschemas(schema, "schema", seen, true);
  const refs = new Map<string, string>();
  const defs: Record<string, unknown> = {};
  const repeated = [...seen.entries()]
    .filter(([, { count, bytes }]) => count > 1 && bytes >= minBytes)
    .toSorted(([, a], [, b]) => b.bytes - a.bytes);
  for (const [key, { name }] of repeated) {
    let defName = name;
    for (let suffix = 2; defName in defs; suffix += 1) defName = `${name}${suffix}`;
    refs.set(key, defName);
    defs[defName] = null;
  }
  if (refs.size === 0) return schema;
  const replace = (node: unknown, isRoot: boolean): unknown => {
    if (Array.isArray(node)) return node.map((item) => replace(item, false));
    if (!isRecord(node)) return node;
    const ref = isRoot ? undefined : refs.get(JSON.stringify(node));
    if (ref !== undefined) return { $ref: `#/$defs/${ref}` };
    const replaced: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(node)) {
      replaced[key] =
        SCHEMA_MAP_KEYWORDS.has(key) && isRecord(value)
          ? Object.fromEntries(
              Object.entries(value).map(([child, sub]) => [child, replace(sub, false)]),
            )
          : replace(value, false);
    }
    return replaced;
  };
  // A hoisted shape can contain a smaller hoisted one: replace inside each definition too.
  for (const [key, defName] of refs) {
    defs[defName] = replace(JSON.parse(key), true);
  }
  const root = replace(schema, true);
  return isRecord(root) ? { ...root, $defs: defs } : schema;
};

/** The UTF-8 size of what `tools/list` sends for these tools. */
export const advertisedBytes = (tools: readonly ListedTool[]): number =>
  new TextEncoder().encode(JSON.stringify(tools)).length;
