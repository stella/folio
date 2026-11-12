import { expect, setDefaultTimeout, test } from "bun:test";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../../../test/subprocess";

setDefaultTimeout(subprocessTestTimeout(2));

// Run after packaging: inspect the shipped archive rather than build inputs.
test("the packed universal VSIX carries the portable backend and no native binaries", () => {
  const result = expectExited(
    runChild(["unzip", "-Z1", path.resolve(import.meta.dir, "../folio.vsix")]),
  );
  expect(result.exitCode).toBe(0);
  const files = result.stdout.trim().split("\n");
  expect(files).toContain("extension/dist/cli/fuzzy-search.wasm32-wasi.wasm");
  expect(files).toContain("extension/dist/cli/folio.mjs");
  expect(files.filter((file) => file.endsWith(".node"))).toEqual([]);
  const wasm = expectExited(
    runChild([
      "unzip",
      "-p",
      path.resolve(import.meta.dir, "../folio.vsix"),
      "extension/dist/cli/fuzzy-search.wasm32-wasi.wasm",
    ]),
  );
  expect(wasm.exitCode).toBe(0);
  // The WebAssembly magic, "\0asm", is ASCII and survives the text decode.
  expect(wasm.stdout.slice(0, 4)).toBe("\0asm");
});
