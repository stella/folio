/**
 * Generate the `ST_Border` union `BorderSpec.style` is typed with.
 *
 * `w:val` on `CT_Border` is `ST_Border`, a 193-member enumeration that the
 * model used to hold as a bare `string` beside a hand-written 22-member
 * `KnownBorderStyle` the parser narrowed against. The hand list omitted every
 * art border and four line styles the schema declares, so a document that used
 * one parsed into a value no consumer's map knew, and the compiler could not
 * say which consumer was missing which member.
 *
 * The enumeration is therefore derived from the committed schema graph rather
 * than restated: a schema refresh that adds or removes a member fails
 * `bun run generate:border-styles:check` instead of silently widening a
 * `string`.
 *
 * Usage:
 *   bun scripts/generate-border-styles.ts write
 *   bun scripts/generate-border-styles.ts check
 */

import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

import { loadSchemaGraph, WML_NAMESPACE } from "./lib/ooxml-schema-graph";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/docx-core/src/model/borderStyle.gen.ts");

const BORDER_SIMPLE_TYPE = `simpleType:{${WML_NAMESPACE}}ST_Border`;

/**
 * Members whose absence would mean the derivation, not the schema, changed.
 *
 * `nil` and `none` are the two reserved members the whole migration turns on,
 * and `apples` is the first art border: a graph that lost the art half of the
 * enumeration would still satisfy a `length > 0` check.
 */
const REQUIRED_MEMBERS = ["nil", "none", "single", "apples", "zigZagStitch", "custom"] as const;

class GenerateBorderStylesError extends TaggedError("GenerateBorderStylesError")<{
  message: string;
}> {}

const renderModule = (members: readonly string[]): string => `/**
 * GENERATED FILE — do not edit.
 *
 * The \`ST_Border\` enumeration, in schema order, derived from
 * \`specifications/generated/docx-transitional-schema.gen.json\` by
 * \`scripts/generate-border-styles.ts\`. Regenerate with:
 *
 *   bun run generate:border-styles
 *
 * The first 27 members are line styles; the rest are the page-border art
 * glyphs \`w:pgBorders\` may name. \`nil\` and \`none\` are distinct: \`none\`
 * cancels a border inherited from the container, \`nil\` states no border. Read
 * either through \`isBorderNone\`/\`isBorderNil\` in \`./borderStyle\`, never by
 * comparing the string.
 */

/** Every \`ST_Border\` member, in schema order. */
export const BORDER_STYLES = [
${members.map((member) => `  "${member}",`).join("\n")}
] as const;

/** \`w:val\` on \`CT_Border\`: a border style the schema declares. */
export type BorderStyle = (typeof BORDER_STYLES)[number];
`;

const main = async (): Promise<void> => {
  const mode = process.argv.at(2) ?? "write";
  if (mode !== "check" && mode !== "write") {
    throw new GenerateBorderStylesError({
      message: "Usage: bun scripts/generate-border-styles.ts [check|write]",
    });
  }

  const graph = await loadSchemaGraph();
  const members = graph.symbols.find((symbol) => symbol.id === BORDER_SIMPLE_TYPE)?.enumValues;
  if (members === undefined || members.length === 0) {
    throw new GenerateBorderStylesError({
      message: `${BORDER_SIMPLE_TYPE} enumerates nothing in the committed schema graph.`,
    });
  }
  for (const required of REQUIRED_MEMBERS) {
    if (!members.includes(required)) {
      throw new GenerateBorderStylesError({
        message: `ST_Border no longer enumerates \`${required}\`; the schema graph or the derivation drifted.`,
      });
    }
  }
  const duplicate = members.find((member, index) => members.indexOf(member) !== index);
  if (duplicate !== undefined) {
    throw new GenerateBorderStylesError({
      message: `ST_Border enumerates \`${duplicate}\` twice; the union would carry a duplicate member.`,
    });
  }

  const rendered = renderModule(members);
  if (mode === "write") {
    await writeFile(OUTPUT_PATH, rendered, "utf8");
    process.stdout.write(`borderStyle.gen.ts written (${String(members.length)} members)\n`);
    return;
  }

  const existing = await readFile(OUTPUT_PATH, "utf8").catch(() => null);
  if (existing !== rendered) {
    throw new GenerateBorderStylesError({
      message: "borderStyle.gen.ts is stale. Run `bun run generate:border-styles`.",
    });
  }
  process.stdout.write("borderStyle.gen.ts is up to date\n");
};

await main();
