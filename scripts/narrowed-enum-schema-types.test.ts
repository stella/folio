/**
 * Every `narrowEnum` value set is the schema enumeration it mirrors.
 *
 * A picklist that has drifted from its simple type is a silent survival loss:
 * `narrowEnum` returns `undefined` for a member it omits and the caller drops
 * the attribute, so the value never reaches the writer. `ThemeColorSlotSchema`
 * drifted that way for sixteen members and nothing noticed, because nothing
 * compared the two.
 *
 * Three assertions, together:
 *   1. the registry is total over the picklists `parserEnums.ts` exports;
 *   2. every `narrowEnum(` call in the packages names a picklist the registry
 *      carries — the wiring, by source scan, so a new reader cannot bypass it;
 *   3. each binding holds, member for member, with a recorded divergence
 *      allowed to shrink but never to grow.
 */

import { readdirSync, readFileSync, statSync } from "node:fs";
import path from "node:path";

import { describe, expect, test } from "bun:test";

import * as parserEnums from "../packages/core/src/docx/parserEnums";

import {
  NARROWED_ENUM_NAMESPACE_URIS,
  NARROWED_ENUM_SCHEMA_TYPES,
  type NarrowedEnumBinding,
} from "./lib/narrowed-enum-schema-types";
import { loadSchemaGraph } from "./lib/ooxml-schema-graph";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const PACKAGES_ROOT = path.join(REPO_ROOT, "packages");

const graph = await loadSchemaGraph();

const enumerationOf = (qualified: string): readonly string[] => {
  const colon = qualified.indexOf(":");
  const prefix = qualified.slice(0, colon);
  const uri = (NARROWED_ENUM_NAMESPACE_URIS as Record<string, string>)[prefix];
  if (uri === undefined) {
    throw new Error(`"${qualified}" uses an undeclared namespace prefix`);
  }
  const name = qualified.slice(colon + 1);
  const symbol = graph.symbols.find(
    (candidate) =>
      candidate.kind === "simpleType" && candidate.name === name && candidate.namespace === uri,
  );
  if (symbol?.enumValues === undefined || symbol.enumValues.length === 0) {
    throw new Error(`"${qualified}" is not an enumerating simple type in the schema graph`);
  }
  return symbol.enumValues;
};

const picklistOptions = (site: string): readonly string[] => {
  const schema = (parserEnums as Record<string, unknown>)[site];
  const options = (schema as { options?: readonly string[] } | undefined)?.options;
  if (options === undefined) {
    throw new Error(`parserEnums.${site} is not a picklist`);
  }
  return options;
};

const PICKLIST_EXPORTS = Object.keys(parserEnums)
  .filter((name) => name.endsWith("Schema"))
  .toSorted();

/** Every TypeScript file under `packages`, tests included. */
const sourceFiles = (dir: string): string[] => {
  const found: string[] = [];
  for (const entry of readdirSync(dir)) {
    if (entry === "node_modules" || entry === "dist") {
      continue;
    }
    const full = path.join(dir, entry);
    if (statSync(full).isDirectory()) {
      found.push(...sourceFiles(full));
      continue;
    }
    if (full.endsWith(".ts") || full.endsWith(".tsx")) {
      found.push(full);
    }
  }
  return found;
};

/**
 * The picklist each `narrowEnum(` call names, by scanning forward from the call
 * to its balanced closing parenthesis. A regular expression cannot do this: the
 * second argument often sits after a nested call spread over several lines.
 */
const narrowEnumArguments = (source: string): string[] => {
  const named: string[] = [];
  const call = /\bnarrowEnum\(/gu;
  let match = call.exec(source);
  while (match !== null) {
    let depth = 1;
    let index = match.index + match[0].length;
    while (index < source.length && depth > 0) {
      const character = source[index];
      if (character === "(") {
        depth += 1;
      } else if (character === ")") {
        depth -= 1;
      }
      index += 1;
    }
    const argumentText = source.slice(match.index + match[0].length, index - 1);
    const identifiers = argumentText.match(/[A-Za-z_$][\w$]*Schema\b/gu) ?? [];
    named.push(...identifiers);
    match = call.exec(source);
  }
  return named;
};

describe("narrowEnum sites are bound to their schema simple type", () => {
  test("the registry is total over the picklists parserEnums exports", () => {
    expect(Object.keys(NARROWED_ENUM_SCHEMA_TYPES).toSorted()).toEqual(PICKLIST_EXPORTS);
  });

  test("every narrowEnum call names a picklist the registry carries", () => {
    const unregistered = new Map<string, string[]>();
    let calls = 0;
    for (const file of sourceFiles(PACKAGES_ROOT)) {
      for (const site of narrowEnumArguments(readFileSync(file, "utf8"))) {
        calls += 1;
        if (site in NARROWED_ENUM_SCHEMA_TYPES) {
          continue;
        }
        const files = unregistered.get(site) ?? [];
        files.push(path.relative(REPO_ROOT, file));
        unregistered.set(site, files);
      }
    }
    // The scan itself has to be known to work; a refactor that renames the
    // helper would otherwise make this test vacuously green.
    expect(calls).toBeGreaterThan(40);
    expect([...unregistered.entries()]).toEqual([]);
  });

  test.each(Object.entries(NARROWED_ENUM_SCHEMA_TYPES) as [string, NarrowedEnumBinding][])(
    "%s",
    (site, binding) => {
      if (binding.kind === "unschematised") {
        expect(binding.reason.length).toBeGreaterThan(0);
        return;
      }

      const options = picklistOptions(site);
      const enumeration = enumerationOf(binding.simpleType);
      const missing = enumeration.filter((value) => !options.includes(value)).toSorted();
      const extra = options.filter((value) => !enumeration.includes(value)).toSorted();

      if (binding.kind === "keeps-raw") {
        // Nothing is lost by narrowing, but an invented member would mean the
        // union accepts a token the format does not declare.
        expect(extra).toEqual([]);
        return;
      }

      expect({ missing, extra }).toEqual(
        binding.kind === "matches"
          ? { missing: [], extra: [] }
          : { missing: [...binding.missing].toSorted(), extra: [...binding.extra].toSorted() },
      );
    },
  );
});
