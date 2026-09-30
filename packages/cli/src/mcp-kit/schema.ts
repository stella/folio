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

const compactNode = (node: unknown): unknown => {
  if (Array.isArray(node)) return node.map(compactNode);
  if (!isRecord(node)) return node;
  const compacted: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(node)) {
    if (ANNOTATION_KEYWORDS.has(key)) continue;
    compacted[key] =
      SCHEMA_MAP_KEYWORDS.has(key) && isRecord(value)
        ? Object.fromEntries(
            Object.entries(value).map(([name, schema]) => [name, compactNode(schema)]),
          )
        : compactNode(value);
  }
  return compacted;
};

/** The schema without annotation keywords, at every depth. */
export const compactSchema = (schema: JsonSchema): JsonSchema => {
  const compacted = compactNode(schema);
  return isRecord(compacted) ? compacted : {};
};

/** The UTF-8 size of what `tools/list` sends for these tools. */
export const advertisedBytes = (tools: readonly ListedTool[]): number =>
  new TextEncoder().encode(JSON.stringify(tools)).length;
