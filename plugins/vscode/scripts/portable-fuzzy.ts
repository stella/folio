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
        const fetchSource = "await globalThis.fetch(__wasmUrl)";
        const poolSource = "const __asyncWorkPoolSize = 4;";
        const reuseSource =
          "reuseWorker: { size: __asyncWorkPoolSize + __workerPoolSize }";
        if (
          ![fetchSource, poolSource, reuseSource].every(
            (expected) => source.split(expected).length === 2,
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
              'new Response(await (await import("node:fs/promises")).readFile(new URL(__wasmUrl)))',
            )
            // The CLI uses synchronous distance calls; it needs no browser workers.
            .replace(poolSource, "const __asyncWorkPoolSize = 0;")
            .replace(reuseSource, "reuseWorker: false"),
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
