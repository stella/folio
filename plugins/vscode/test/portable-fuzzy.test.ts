import { expect, test } from "bun:test";
import { build } from "esbuild";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { copyFuzzyWasm, fuzzyFromDisk } from "../scripts/portable-fuzzy";

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
        js: "import { createRequire } from 'node:module'; const require = createRequire(import.meta.url);",
      },
      plugins: [fuzzyFromDisk],
    });
    expect(Object.keys(result.metafile.inputs).some((file) => file.endsWith(".node"))).toBe(false);
    await copyFuzzyWasm(directory);
    const child = Bun.spawnSync(["node", outfile]);
    expect(child.exitCode).toBe(0);
    expect(JSON.parse(child.stdout.toString())).toEqual([3, 2, 1, 1, 0]);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
