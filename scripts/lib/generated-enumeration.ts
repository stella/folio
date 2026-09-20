/**
 * Rendering and write/check plumbing shared by the schema-derived enumeration
 * generators.
 *
 * Every one of them answers the same question — which tokens does this simple
 * type enumerate — and emits the same shape: a frozen array and a spelled-out
 * union, from one member list, so the two cannot drift. Sharing the renderer
 * means a change to that shape lands in one place rather than once per
 * generator.
 */

import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";

import { TaggedError } from "better-result";

import type { Index } from "./ooxml-schema-graph";

export class GeneratedEnumerationError extends TaggedError("GeneratedEnumerationError")<{
  message: string;
}> {}

/** The tokens one enumerating simple type declares, in schema order. */
export const enumerationOf = (index: Index, namespace: string, name: string): readonly string[] => {
  const symbol = index.byId.get(`simpleType:{${namespace}}${name}`);
  if (symbol?.enumValues === undefined || symbol.enumValues.length === 0) {
    throw new GeneratedEnumerationError({
      message: `${name} is not an enumerating simple type in the committed schema graph.`,
    });
  }
  return symbol.enumValues;
};

export type RenderListOptions = {
  /** The array's exported name. */
  name: string;
  /** The union's exported name. */
  type: string;
  /** The doc comment both carry, already spelled as a block comment. */
  doc: string;
  members: readonly string[];
};

/**
 * The array and the union, emitted together from one member list.
 *
 * The union is spelled out rather than written `(typeof NAME)[number]`, so a
 * published declaration does not depend on an array the package has no reason
 * to export.
 */
export const renderList = ({ name, type, doc, members }: RenderListOptions): string =>
  `${doc}
export const ${name} = [
${members.map((member) => `  "${member}",`).join("\n")}
] as const;

export type ${type} =
${members.map((member) => `  | "${member}"`).join("\n")};
`;

export type RenderMapOptions = {
  /** The map's exported name. */
  name: string;
  /** The union the keys come from. */
  keyType: string;
  /** The union the values come from. */
  valueType: string;
  /** The doc comment, already spelled as a block comment. */
  doc: string;
  /** Key to value, in the order the keys are to be emitted. */
  entries: readonly (readonly [string, string])[];
};

/**
 * A total map over one enumeration, emitted as `satisfies Record<K, V>`.
 *
 * The `satisfies` is what makes it total: a key the enumeration gains and the
 * map does not fails to compile, which is the only thing that keeps a
 * hand-decided table from drifting away from the generated union beside it.
 */
export const renderMap = ({ name, keyType, valueType, doc, entries }: RenderMapOptions): string =>
  `${doc}
export const ${name} = {
${entries.map(([key, value]) => `  ${key}: "${value}",`).join("\n")}
} as const satisfies Record<${keyType}, ${valueType}>;
`;

export type RenderModuleOptions = {
  /** What the module holds, as the generated header's first paragraph. */
  summary: string;
  /** The `package.json` script that rewrites it. */
  script: string;
  lists: readonly string[];
};

export const renderModule = ({ summary, script, lists }: RenderModuleOptions): string => `/**
 * GENERATED FILE — do not edit.
 *
${summary
  .split("\n")
  .map((line) => (line === "" ? " *" : ` * ${line}`))
  .join("\n")}
 *
 * Derived from \`specifications/generated/docx-transitional-schema.gen.json\`.
 * Regenerate with:
 *
 *   bun run ${script}
 */

${lists.join("\n")}`;

const digest = (value: string): string => createHash("sha256").update(value).digest("hex");

export type EmitOptions = {
  /** `"write"` rewrites the module; `"check"` fails when it is stale. */
  mode: string;
  outputPath: string;
  rendered: string;
  /** What `write` prints, so a regeneration says what it produced. */
  summary: string;
  /** The `package.json` script `check` names when the module is stale. */
  script: string;
};

/** Write the module, or fail when the committed one differs from it. */
export const emitGeneratedModule = async ({
  mode,
  outputPath,
  rendered,
  summary,
  script,
}: EmitOptions): Promise<void> => {
  if (mode !== "check" && mode !== "write") {
    throw new GeneratedEnumerationError({
      message: `Usage: bun <generator> [check|write], not "${mode}".`,
    });
  }
  const name = outputPath.slice(outputPath.lastIndexOf("/") + 1);
  if (mode === "write") {
    await writeFile(outputPath, rendered, "utf8");
    process.stdout.write(`${name} written (${summary})\n`);
    return;
  }
  const existing = await readFile(outputPath, "utf8").catch(() => null);
  if (existing === null || digest(existing) !== digest(rendered)) {
    throw new GeneratedEnumerationError({
      message: `${name} is stale. Run \`bun run ${script}\`.`,
    });
  }
  process.stdout.write(`${name} is up to date\n`);
};
