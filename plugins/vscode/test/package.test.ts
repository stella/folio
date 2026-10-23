import { expect, test } from "bun:test";
import path from "node:path";

// Run after packaging: inspect the shipped archive rather than build inputs.
test("the packed universal VSIX carries the portable backend and no native binaries", () => {
  const result = Bun.spawnSync([
    "unzip",
    "-Z1",
    path.resolve(import.meta.dir, "../folio.vsix"),
  ]);
  expect(result.exitCode).toBe(0);
  const files = result.stdout.toString().trim().split("\n");
  expect(files).toContain("extension/dist/cli/fuzzy-search.wasm32-wasi.wasm");
  expect(files).toContain("extension/dist/cli/folio.mjs");
  expect(files.filter((file) => file.endsWith(".node"))).toEqual([]);
  const wasm = Bun.spawnSync([
    "unzip",
    "-p",
    path.resolve(import.meta.dir, "../folio.vsix"),
    "extension/dist/cli/fuzzy-search.wasm32-wasi.wasm",
  ]);
  expect(wasm.exitCode).toBe(0);
  expect([...wasm.stdout.subarray(0, 4)]).toEqual([0, 97, 115, 109]);
});
