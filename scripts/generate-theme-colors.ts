/**
 * Generate the four colour enumerations a theme reference passes through.
 *
 * A `w:themeColor` attribute and an `a:clrScheme` slot name the same colour in
 * two vocabularies, and the two enumerations do not share a single spelling of
 * the hyperlink slot: WordprocessingML writes `hyperlink`, DrawingML writes
 * `hlink`. A hand-written union for either one drifts silently, because nothing
 * binds it to the schema; a token the union omits is dropped at parse time and
 * the attribute never reaches the writer.
 *
 * So all four come from the committed schema graph:
 *
 *   ThemeColor          ST_ThemeColor            `w:themeColor`, `w:themeFill`
 *   SchemeColorSlot     CT_ColorScheme children  `a:clrScheme`'s slot elements
 *   SchemeColorValue    ST_SchemeColorVal        `a:schemeClr/@val`
 *   ClrSchemeMappingKey CT_ColorSchemeMapping    `w:clrSchemeMapping` attributes
 *
 * `SchemeColorSlot` is derived from the content model rather than from
 * `ST_ColorSchemeIndex`, because the slots a theme part declares are what a
 * lookup can hit; the generator asserts the two agree, so a schema refresh that
 * separates them fails here instead of in a painter.
 *
 * Usage:
 *   bun scripts/generate-theme-colors.ts write
 *   bun scripts/generate-theme-colors.ts check
 */

import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

import {
  buildIndex,
  type Index,
  loadSchemaGraph,
  type SchemaGraph,
} from "./lib/ooxml-schema-graph";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/docx-core/src/model/themeColor.gen.ts");

const WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DML = "http://schemas.openxmlformats.org/drawingml/2006/main";

class GenerateThemeColorsError extends TaggedError("GenerateThemeColorsError")<{
  message: string;
}> {}

const enumerationOf = (index: Index, namespace: string, name: string): readonly string[] => {
  const symbol = index.byId.get(`simpleType:{${namespace}}${name}`);
  if (symbol?.enumValues === undefined || symbol.enumValues.length === 0) {
    throw new GenerateThemeColorsError({
      message: `${name} is not an enumerating simple type in the committed schema graph.`,
    });
  }
  return symbol.enumValues;
};

/**
 * The slot elements `a:clrScheme` declares, in schema order.
 *
 * `a:extLst` is the extension list every DrawingML container carries, not a
 * colour slot.
 */
const colorSchemeSlots = (graph: SchemaGraph): readonly string[] => {
  const owner = `complexType:{${DML}}CT_ColorScheme`;
  const slots = graph.children
    .filter((child) => child.owner === owner && child.kind === "element" && child.name !== "extLst")
    .toSorted((left, right) => (left.order ?? 0) - (right.order ?? 0))
    .map((child) => child.name)
    .filter((name): name is string => name !== undefined);
  if (slots.length === 0) {
    throw new GenerateThemeColorsError({
      message: "CT_ColorScheme declares no slot elements; the schema graph drifted.",
    });
  }
  return slots;
};

/** The attribute names `w:clrSchemeMapping` carries, in schema order. */
const clrSchemeMappingKeys = (graph: SchemaGraph): readonly string[] => {
  const owner = `complexType:{${WML}}CT_ColorSchemeMapping`;
  const keys = graph.attributes
    .filter((attribute) => attribute.owner === owner && attribute.kind === "attribute")
    .toSorted((left, right) => (left.order ?? 0) - (right.order ?? 0))
    .map((attribute) => attribute.name)
    .filter((name): name is string => name !== undefined);
  if (keys.length === 0) {
    throw new GenerateThemeColorsError({
      message: "CT_ColorSchemeMapping declares no attributes; the schema graph drifted.",
    });
  }
  return keys;
};

const sameMembers = (left: readonly string[], right: readonly string[]): boolean =>
  left.length === right.length && [...left].toSorted().join() === [...right].toSorted().join();

/**
 * The array and the union, emitted together from one member list.
 *
 * The union is spelled out rather than written `(typeof NAME)[number]`, so a
 * published declaration does not depend on an array the package has no reason
 * to export. One call emits both, so the two cannot drift.
 */
const renderList = (name: string, type: string, doc: string, members: readonly string[]): string =>
  `${doc}
export const ${name} = [
${members.map((member) => `  "${member}",`).join("\n")}
] as const;

export type ${type} =
${members.map((member) => `  | "${member}"`).join("\n")};
`;

const renderModule = (lists: readonly string[]): string => `/**
 * GENERATED FILE — do not edit.
 *
 * The colour enumerations a theme reference passes through, derived from
 * \`specifications/generated/docx-transitional-schema.gen.json\` by
 * \`scripts/generate-theme-colors.ts\`. Regenerate with:
 *
 *   bun run generate:theme-colors
 */

${lists.join("\n")}`;

const digest = (value: string): string => createHash("sha256").update(value).digest("hex");

const main = async (): Promise<void> => {
  const mode = process.argv.at(2) ?? "write";
  if (mode !== "check" && mode !== "write") {
    throw new GenerateThemeColorsError({
      message: "Usage: bun scripts/generate-theme-colors.ts [check|write]",
    });
  }

  const graph = await loadSchemaGraph();
  const index = buildIndex(graph);

  const themeColors = enumerationOf(index, WML, "ST_ThemeColor");
  const schemeColorValues = enumerationOf(index, DML, "ST_SchemeColorVal");
  const schemeColorSlots = colorSchemeSlots(graph);
  const mappingKeys = clrSchemeMappingKeys(graph);

  const schemeColorIndex = enumerationOf(index, DML, "ST_ColorSchemeIndex");
  if (!sameMembers(schemeColorSlots, schemeColorIndex)) {
    throw new GenerateThemeColorsError({
      message:
        `CT_ColorScheme declares ${schemeColorSlots.join(", ")} but ST_ColorSchemeIndex ` +
        `enumerates ${schemeColorIndex.join(", ")}. A slot a theme part can declare and a slot ` +
        "a mapping can name must stay the same set; decide which one the model follows.",
    });
  }

  const rendered = renderModule([
    renderList(
      "THEME_COLORS",
      "ThemeColor",
      `/**
 * \`ST_ThemeColor\`: every token a WordprocessingML \`w:themeColor\`,
 * \`w:themeFill\` or \`w:clrSchemeMapping\` value may carry.
 *
 * Spelled in full words (\`dark1\`, \`hyperlink\`), unlike the DrawingML slot
 * names below. \`none\` is a reserved member that cancels an inherited theme
 * colour rather than naming one.
 */`,
      themeColors,
    ),
    renderList(
      "SCHEME_COLOR_SLOTS",
      "SchemeColorSlot",
      `/**
 * The colour slots a theme part's \`a:clrScheme\` declares, in schema order.
 *
 * This is the set a theme lookup can hit; {@link SCHEME_COLOR_VALUES} is wider.
 */`,
      schemeColorSlots,
    ),
    renderList(
      "SCHEME_COLOR_VALUES",
      "SchemeColorValue",
      `/**
 * \`ST_SchemeColorVal\`: every token an \`a:schemeClr/@val\` reference may carry.
 *
 * A superset of {@link SCHEME_COLOR_SLOTS}: it adds the four mapped spellings
 * (\`bg1\`, \`tx1\`, \`bg2\`, \`tx2\`) and \`phClr\`, the placeholder a style
 * definition resolves against its instantiating context.
 */`,
      schemeColorValues,
    ),
    renderList(
      "CLR_SCHEME_MAPPING_KEYS",
      "ClrSchemeMappingKey",
      `/**
 * The attributes \`w:clrSchemeMapping\` carries in \`settings.xml\`, each naming
 * the theme slot one mapped colour resolves to.
 */`,
      mappingKeys,
    ),
  ]);

  if (mode === "write") {
    await writeFile(OUTPUT_PATH, rendered, "utf8");
    process.stdout.write(
      `themeColor.gen.ts written (${String(themeColors.length)} theme colours, ` +
        `${String(schemeColorSlots.length)} scheme slots, ` +
        `${String(schemeColorValues.length)} scheme values, ` +
        `${String(mappingKeys.length)} mapping keys)\n`,
    );
    return;
  }

  const existing = await readFile(OUTPUT_PATH, "utf8").catch(() => null);
  if (existing === null || digest(existing) !== digest(rendered)) {
    throw new GenerateThemeColorsError({
      message: "themeColor.gen.ts is stale. Run `bun run generate:theme-colors`.",
    });
  }
  process.stdout.write("themeColor.gen.ts is up to date\n");
};

await main();
