import type { Plugin } from "esbuild";
import { copyFile, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const FUZZY_WASM = "fuzzy-search.wasm32-wasi.wasm";
const fuzzyEntry = fileURLToPath(
  import.meta.resolve("@stll/fuzzy-search-wasm"),
);

/** Use the portable backend in the Node.js CLI, reading its asset from disk. */
export const fuzzyFromDisk: Plugin = {
  name: "folio-fuzzy-from-disk",
  setup: (builder) => {
    builder.onResolve({ filter: /^@stll\/fuzzy-search$/ }, () => ({
      path: fuzzyEntry,
    }));
    builder.onLoad(
      { filter: /[\\/]dist[\\/]wasm\.mjs$/ },
      async ({ path: file }) => {
        if (file !== fuzzyEntry) return;
        const source = await readFile(file, "utf8");
        const fetchSource = "await fetch(__wasmUrl).then(async (res) => {";
        const poolSource = "asyncWorkPoolSize: 4";
        const bytesSource = "const bytes = await res.arrayBuffer();";
        if (
          ![fetchSource, poolSource, bytesSource].every((expected) =>
            source.includes(expected),
          )
        ) {
          throw new Error(
            `${file} no longer initializes its wasm as this build expects.`,
          );
        }
        return {
          contents: source
            .replace(
              fetchSource,
              'await import("node:fs/promises").then(async (fs) => {',
            )
            .replace(
              bytesSource,
              "const bytes = await fs.readFile(new URL(__wasmUrl));",
            )
            // The CLI uses synchronous distance calls; it needs no browser workers.
            .replace(poolSource, "asyncWorkPoolSize: 0"),
          loader: "js",
        };
      },
    );
  },
};

export const copyFuzzyWasm = async (directory: string): Promise<void> => {
  await copyFile(
    path.join(path.dirname(fuzzyEntry), FUZZY_WASM),
    path.join(directory, FUZZY_WASM),
  );
};
