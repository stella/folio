#!/usr/bin/env bun
// Write the hyphenation dictionaries @stll/folio-core fetches at runtime:
// `packages/core/assets/hyphenation/<id>.json`, the pattern data of the pinned
// `hyphen` package, plus that package's licence. They ship as data, not
// modules, so no bundle of folio carries a pattern trie.
//
// Usage: `bun scripts/hyphenation-assets-gen.ts` writes them; `--check` exits
// non-zero when a written file differs from what the package gives.

import { mkdir, readFile, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";

import { HYPHENATION_DICTIONARY_IDS } from "../packages/core/src/layout-engine/measure/hyphenationDictionaries";

const coreDir = path.resolve(import.meta.dir, "../packages/core");
const outDir = path.join(coreDir, "assets/hyphenation");
const require = createRequire(path.join(coreDir, "package.json"));

const expected = async (): Promise<Map<string, string>> => {
  const files = new Map<string, string>();
  for (const id of HYPHENATION_DICTIONARY_IDS) {
    const patterns: unknown = require(`hyphen/patterns/${id}.js`);
    files.set(`${id}.json`, `${JSON.stringify(patterns)}\n`);
  }
  files.set(
    "LICENSE",
    await readFile(
      path.join(path.dirname(require.resolve("hyphen")), "LICENSE"),
      "utf8",
    ),
  );
  return files;
};

const files = await expected();
if (process.argv.includes("--check")) {
  const stale: string[] = [];
  for (const [name, contents] of files) {
    const written = await readFile(path.join(outDir, name), "utf8").catch(
      () => undefined,
    );
    if (written !== contents) stale.push(name);
  }
  if (stale.length > 0) {
    console.error(
      `Stale hyphenation assets (${stale.join(", ")}); run bun scripts/hyphenation-assets-gen.ts`,
    );
    process.exit(1);
  }
} else {
  await mkdir(outDir, { recursive: true });
  for (const [name, contents] of files) {
    await writeFile(path.join(outDir, name), contents);
  }
}
