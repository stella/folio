/**
 * Generate the `ST_PresetLineDashVal` union a shape or text-box outline is typed with.
 *
 * `a:prstDash@val` is DrawingML's dash vocabulary. The model restated it as a
 * hand-written eleven-member union, and the painter resolved it through a table
 * keyed by lower-cased strings shared with CSS `border-style` and CSS
 * `text-decoration-style`, so `dash`, `dot` and `solid` collided across three
 * vocabularies and the eight members no vocabulary spelled the same way fell
 * through to a plain line.
 *
 * The enumeration is derived from the committed schema graph rather than
 * restated: a schema refresh that adds or removes a member fails
 * `bun run generate:preset-line-dash:check` instead of leaving one map short.
 *
 * Usage:
 *   bun scripts/generate-preset-line-dash.ts write
 *   bun scripts/generate-preset-line-dash.ts check
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

import { loadSchemaGraph } from "./lib/ooxml-schema-graph";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/docx-core/src/model/presetLineDash.gen.ts");

const DRAWINGML_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PRESET_DASH_SIMPLE_TYPE = `simpleType:{${DRAWINGML_NAMESPACE}}ST_PresetLineDashVal`;

/**
 * Members whose absence would mean the derivation, not the schema, changed.
 *
 * `solid` is the one member that paints an unbroken line, `dot` and `dash` are
 * the two the old shared table happened to know, and `sysDashDotDot` is the
 * last of the system patterns: a graph that lost the `sys*` half would still
 * satisfy a `length > 0` check.
 */
const REQUIRED_MEMBERS = ["solid", "dot", "dash", "sysDashDotDot"] as const;

class GeneratePresetLineDashError extends TaggedError("GeneratePresetLineDashError")<{
  message: string;
}> {}

const renderModule = (members: readonly string[]): string => `/**
 * GENERATED FILE — do not edit.
 *
 * The \`ST_PresetLineDashVal\` enumeration, in schema order, derived from
 * \`specifications/generated/docx-transitional-schema.gen.json\` by
 * \`scripts/generate-preset-line-dash.ts\`. Regenerate with:
 *
 *   bun run generate:preset-line-dash
 *
 * It is the dash vocabulary of \`a:ln/a:prstDash@val\`, and only that: the
 * spellings it shares with CSS \`border-style\` (\`solid\`, \`dash\`, \`dot\`)
 * name different patterns there. Read it through \`./presetLineDash\`.
 */

/** Every \`ST_PresetLineDashVal\` member, in schema order. */
export const PRESET_LINE_DASH_VALS = [
${members.map((member) => `  "${member}",`).join("\n")}
] as const;

/** \`a:prstDash@val\`: a preset dash the schema declares. */
export type PresetLineDashVal = (typeof PRESET_LINE_DASH_VALS)[number];
`;

const main = async (): Promise<void> => {
  const mode = process.argv.at(2) ?? "write";
  if (mode !== "check" && mode !== "write") {
    throw new GeneratePresetLineDashError({
      message: "Usage: bun scripts/generate-preset-line-dash.ts [check|write]",
    });
  }

  const graph = await loadSchemaGraph();
  const members = graph.symbols.find((symbol) => symbol.id === PRESET_DASH_SIMPLE_TYPE)?.enumValues;
  if (members === undefined || members.length === 0) {
    throw new GeneratePresetLineDashError({
      message: `${PRESET_DASH_SIMPLE_TYPE} enumerates nothing in the committed schema graph.`,
    });
  }
  for (const required of REQUIRED_MEMBERS) {
    if (!members.includes(required)) {
      throw new GeneratePresetLineDashError({
        message: `ST_PresetLineDashVal no longer enumerates \`${required}\`; the schema graph or the derivation drifted.`,
      });
    }
  }
  const duplicate = members.find((member, index) => members.indexOf(member) !== index);
  if (duplicate !== undefined) {
    throw new GeneratePresetLineDashError({
      message: `ST_PresetLineDashVal enumerates \`${duplicate}\` twice; the union would carry a duplicate member.`,
    });
  }

  const rendered = renderModule(members);
  if (mode === "write") {
    await writeFile(OUTPUT_PATH, rendered, "utf8");
    process.stdout.write(`presetLineDash.gen.ts written (${String(members.length)} members)\n`);
    return;
  }

  const existing = await readFile(OUTPUT_PATH, "utf8").catch(() => null);
  if (existing !== rendered) {
    throw new GeneratePresetLineDashError({
      message: "presetLineDash.gen.ts is stale. Run `bun run generate:preset-line-dash`.",
    });
  }
  process.stdout.write("presetLineDash.gen.ts is up to date\n");
};

await main();
