import { expect, setDefaultTimeout, test } from "bun:test";
import { build } from "esbuild";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../../../test/subprocess";
import { copyFuzzyWasm, fuzzyFromDisk } from "../scripts/portable-fuzzy";

setDefaultTimeout(subprocessTestTimeout());

test("the universal CLI backend loads from disk and computes Unicode distances in Node", async () => {
  const directory = await mkdtemp(path.join(tmpdir(), "folio-fuzzy-"));
  try {
    const outfile = path.join(directory, "distance.mjs");
    const result = await build({
      stdin: {
        contents: `import { distance } from "@stll/fuzzy-search";
          console.log(JSON.stringify([
            distance("kitten", "sitting"),
            distance("ab", "ba", "levenshtein"),
            distance("ab", "ba", "damerau-levenshtein"),
            distance("😀", "😁"), distance("", ""),
          ]));`,
        resolveDir: path.resolve(import.meta.dir, ".."),
      },
      outfile,
      bundle: true,
      platform: "node",
      format: "esm",
      metafile: true,
      banner: {
        js: `import { createRequire } from 'node:module';
          const require = createRequire(import.meta.url);
          globalThis.fetch = () => { throw new Error("The CLI must read its wasm from disk"); };
          globalThis.Worker = class { constructor() { throw new Error("Synchronous distances need no workers"); } };`,
      },
      plugins: [fuzzyFromDisk],
    });
    expect(
      Object.keys(result.metafile.inputs).some((file) =>
        file.endsWith(".node"),
      ),
    ).toBe(false);
    await copyFuzzyWasm(directory);
    const child = expectExited(runChild(["node", outfile]));
    expect(child.exitCode).toBe(0);
    expect(JSON.parse(child.stdout)).toEqual([3, 2, 1, 1, 0]);
    await writeFile(
      path.join(directory, "fuzzy-search.wasm32-wasi.wasm"),
      "<html>missing wasm</html>",
    );
    const corrupt = expectExited(runChild(["node", outfile]));
    expect(corrupt.exitCode).not.toBe(0);
    expect(corrupt.stderr).toContain("did not contain WebAssembly bytes");
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
