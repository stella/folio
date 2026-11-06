/**
 * The extensions entry is public, and an importer may reach it before anything
 * else: building the singleton schema during that import would run the starter
 * kit before its own module has finished loading.
 */

import { expect, test } from "bun:test";
import { resolve } from "node:path";

test("the extensions entry loads when it is the first thing imported", () => {
  const entry = resolve(import.meta.dir, "index.ts");
  const run = Bun.spawnSync([
    process.execPath,
    "-e",
    `const { createStarterKit } = await import(${JSON.stringify(entry)});
     if (typeof createStarterKit !== "function") process.exit(2);`,
  ]);
  expect(run.stderr.toString()).toBe("");
  expect(run.exitCode).toBe(0);
});
