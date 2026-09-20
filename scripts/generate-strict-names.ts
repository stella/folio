/**
 * Generate the table that says which WordprocessingML name Part 1 renamed.
 *
 * ECMA-376 Part 1 spells a handful of slots by writing direction where Part 4
 * spells them by physical side — `w:start` for `w:left` on a border, `@w:end`
 * for `@w:right` on an indent. folio rebuilds every package as Transitional, so
 * every reader has to take the logical spelling and every writer has to emit
 * the physical one, and the survival census has to know the two are one slot or
 * it reads folio's own canonical output as a loss.
 *
 * The repository vendors no Strict schema, so the pair's *direction* comes from
 * the cited list in `specifications/strict-names/renames.ts`. Everything else is
 * derived: this script finds every slot in the committed Transitional graph
 * where both spellings are declared on one owner with one type, which is what
 * Part 4 does for a name Part 1 renamed. A rename the graph never realises
 * fails here rather than shipping an equivalence nothing backs.
 *
 * Usage:
 *   bun scripts/generate-strict-names.ts write
 *   bun scripts/generate-strict-names.ts check
 */

import { createHash } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

import sourceManifest from "../specifications/sources.json";
import { STRICT_NAME_RENAMES } from "../specifications/strict-names/renames";
import {
  attributesOf,
  buildIndex,
  type Index,
  localName,
  loadSchemaGraph,
  type SchemaGraph,
  WML_NAMESPACE,
} from "./lib/ooxml-schema-graph";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const OUTPUT_PATH = path.join(REPO_ROOT, "packages/core/src/docx/strictNames.gen.ts");

class GenerateStrictNamesError extends TaggedError("GenerateStrictNamesError")<{
  message: string;
}> {}

/** One slot the graph declares under both spellings. */
type Site = {
  kind: "attribute" | "element";
  /** The complex type the renamed thing belongs to: the element's own, or the attribute's owner. */
  type: string;
  strict: string;
  transitional: string;
  /** The complex type that declares the pair, for the comment the module carries. */
  owner: string;
};

const wmlLocalName = (qualifiedName: string | undefined): string | undefined =>
  qualifiedName === undefined || !qualifiedName.startsWith(`{${WML_NAMESPACE}}`)
    ? undefined
    : localName(qualifiedName);

/**
 * Element children each complex type declares, by local name, with their types.
 *
 * A child reached through a `ref` carries its type on the global declaration,
 * so both forms resolve here rather than at each caller.
 */
const wmlChildTypes = (graph: SchemaGraph, index: Index): Map<string, Map<string, string>> => {
  const byOwner = new Map<string, Map<string, string>>();
  for (const child of graph.children) {
    if (child.kind !== "element") {
      continue;
    }
    const declared = child.ref ? index.byId.get(`element:${child.ref}`) : undefined;
    const namespace = declared?.namespace ?? child.namespace;
    const name = declared?.name ?? child.name;
    const type = wmlLocalName(declared?.type ?? child.type);
    if (namespace !== WML_NAMESPACE || name === undefined || type === undefined) {
      continue;
    }
    const children = byOwner.get(child.owner) ?? new Map<string, string>();
    children.set(name, type);
    byOwner.set(child.owner, children);
  }
  return byOwner;
};

const collectSites = (graph: SchemaGraph, index: Index): Site[] => {
  const sites: Site[] = [];
  const bothSpellings = <Value>(
    declared: ReadonlyMap<string, Value>,
    record: (
      rename: (typeof STRICT_NAME_RENAMES)[number],
      strict: Value,
      transitional: Value,
    ) => void,
  ): void => {
    for (const rename of STRICT_NAME_RENAMES) {
      const strict = declared.get(rename.strict);
      const transitional = declared.get(rename.transitional);
      if (strict !== undefined && transitional !== undefined) {
        record(rename, strict, transitional);
      }
    }
  };

  for (const [owner, children] of wmlChildTypes(graph, index)) {
    const ownerName = wmlLocalName(owner.replace(/^complexType:/u, ""));
    if (ownerName === undefined) {
      continue;
    }
    bothSpellings(children, (rename, strictType, transitionalType) => {
      if (strictType !== transitionalType) {
        throw new GenerateStrictNamesError({
          message: `${ownerName} declares w:${rename.strict} as ${strictType} and w:${rename.transitional} as ${transitionalType}; two types are not one slot under two names.`,
        });
      }
      sites.push({
        kind: "element",
        owner: ownerName,
        strict: rename.strict,
        transitional: rename.transitional,
        type: strictType,
      });
    });
  }

  for (const symbol of graph.symbols) {
    if (symbol.kind !== "complexType" || symbol.namespace !== WML_NAMESPACE) {
      continue;
    }
    const declared = new Map(
      attributesOf(index, symbol.id).map((attribute) => [attribute.name, attribute.type]),
    );
    bothSpellings(declared, (rename, strictType, transitionalType) => {
      if (strictType !== transitionalType) {
        throw new GenerateStrictNamesError({
          message: `${symbol.name} types @w:${rename.strict} as ${localName(strictType)} and @w:${rename.transitional} as ${localName(transitionalType)}; two types are not one slot under two names.`,
        });
      }
      sites.push({
        kind: "attribute",
        owner: symbol.name,
        strict: rename.strict,
        transitional: rename.transitional,
        type: symbol.name,
      });
    });
  }

  return sites;
};

/** The key both tables use: the type, then the name, marked `@` for an attribute. */
const slotKey = (type: string, kind: Site["kind"], name: string): string =>
  `${type} ${kind === "attribute" ? "@" : ""}${name}`;

type Entry = {
  key: string;
  owners: string[];
  strict: string;
  transitional: string;
  transitionalKey: string;
};

const collectEntries = (sites: readonly Site[]): Entry[] => {
  const byKey = new Map<string, Entry>();
  for (const site of sites) {
    const key = slotKey(site.type, site.kind, site.strict);
    const existing = byKey.get(key);
    if (existing === undefined) {
      byKey.set(key, {
        key,
        owners: [site.owner],
        strict: site.strict,
        transitional: site.transitional,
        transitionalKey: slotKey(site.type, site.kind, site.transitional),
      });
      continue;
    }
    if (existing.transitional !== site.transitional) {
      throw new GenerateStrictNamesError({
        message: `${key} is renamed to both ${existing.transitional} and ${site.transitional}; the slot key cannot say which applies.`,
      });
    }
    if (!existing.owners.includes(site.owner)) {
      existing.owners.push(site.owner);
    }
  }

  for (const rename of STRICT_NAME_RENAMES) {
    if (![...byKey.values()].some((entry) => entry.strict === rename.strict)) {
      throw new GenerateStrictNamesError({
        message: `no type in the committed graph declares both w:${rename.strict} and w:${rename.transitional}, so the rename has no Transitional side to stand for. Drop it, or refresh the schema graph.`,
      });
    }
  }

  return [...byKey.values()].toSorted((left, right) => (left.key < right.key ? -1 : 1));
};

/** A rename's direction is only as good as the source it cites. */
const checkCitations = (): void => {
  for (const { strict, source } of STRICT_NAME_RENAMES) {
    const cited = sourceManifest.sources.find(({ id }) => id === source);
    if (cited === undefined) {
      throw new GenerateStrictNamesError({
        message: `${strict} cites ${source}, which is not a source in specifications/sources.json.`,
      });
    }
    if (!cited.profiles.includes("strict")) {
      throw new GenerateStrictNamesError({
        message: `${strict} cites ${source}, whose profiles do not include "strict", so it cannot say which spelling Part 1 declares.`,
      });
    }
  }
};

const renderModule = (entries: readonly Entry[]): string => {
  const strictNames = entries.map(
    ({ key, owners, transitional }) =>
      `  // declared by ${owners.join(", ")}\n  "${key}": "${transitional}",`,
  );
  const byTransitional = new Map<string, string[]>();
  for (const { strict, transitionalKey } of entries) {
    byTransitional.set(transitionalKey, [...(byTransitional.get(transitionalKey) ?? []), strict]);
  }
  const transitionalNames = [...byTransitional.entries()]
    .toSorted(([left], [right]) => (left < right ? -1 : 1))
    .map(([key, spellings]) => `  "${key}": [${spellings.map((name) => `"${name}"`).join(", ")}],`);

  return `/**
 * GENERATED FILE — do not edit.
 *
 * The WordprocessingML names ECMA-376 Part 1 spells by writing direction and
 * Part 4 spells by physical side, derived from the cited list in
 * \`specifications/strict-names/renames.ts\` and the committed schema graph by
 * \`scripts/generate-strict-names.ts\`. Regenerate with:
 *
 *   bun run generate:strict-names
 *
 * A key is the complex type the name belongs to — the element's own type, or
 * the attribute's owner — then the local name, marked \`@\` for an attribute.
 * The comment above each entry names the types that declare the pair.
 */

/** What folio writes for a name a Strict producer spelled by writing direction. */
export const TRANSITIONAL_NAME_BY_STRICT_NAME = {
${strictNames.join("\n")}
} as const;

/** A name Part 1 renamed, keyed by the complex type it belongs to. */
export type StrictName = keyof typeof TRANSITIONAL_NAME_BY_STRICT_NAME;

/** The same keys as a list, so a consumer can walk them without widening them. */
export const STRICT_NAMES = [
${entries.map(({ key }) => `  "${key}",`).join("\n")}
] as const satisfies readonly StrictName[];

/** What a Strict producer may have written where folio writes the Transitional name. */
export const STRICT_NAMES_BY_TRANSITIONAL_NAME = {
${transitionalNames.join("\n")}
} as const;

/** A slot a reader must take in either spelling, keyed the same way. */
export type RenamedSlot = keyof typeof STRICT_NAMES_BY_TRANSITIONAL_NAME;
`;
};

const digest = (value: string): string => createHash("sha256").update(value).digest("hex");

const main = async (): Promise<void> => {
  const mode = process.argv.at(2) ?? "write";
  if (mode !== "check" && mode !== "write") {
    throw new GenerateStrictNamesError({
      message: "Usage: bun scripts/generate-strict-names.ts [check|write]",
    });
  }

  checkCitations();
  const graph = await loadSchemaGraph();
  const entries = collectEntries(collectSites(graph, buildIndex(graph)));

  // The two slots every consumer of this table reads: a border side, which is
  // an element rename, and an indent edge, which is an attribute rename. A
  // derivation that stops finding either has drifted, whatever else it found.
  for (const key of ["CT_Border start", "CT_Ind @start"]) {
    if (entries.find((entry) => entry.key === key)?.transitional !== "left") {
      throw new GenerateStrictNamesError({
        message: `${key} must be renamed to left; the schema graph or the derivation drifted.`,
      });
    }
  }

  const rendered = renderModule(entries);
  if (mode === "write") {
    await writeFile(OUTPUT_PATH, rendered, "utf8");
    process.stdout.write(`strictNames.gen.ts written (${entries.length} renamed slots)\n`);
    return;
  }

  const existing = await readFile(OUTPUT_PATH, "utf8").catch(() => null);
  if (existing === null || digest(existing) !== digest(rendered)) {
    throw new GenerateStrictNamesError({
      message: "strictNames.gen.ts is stale. Run `bun run generate:strict-names`.",
    });
  }
  process.stdout.write("strictNames.gen.ts is up to date\n");
};

await main();
