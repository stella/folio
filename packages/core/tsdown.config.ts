import { defineConfig } from "tsdown";

import { stableRelativeImportOrder } from "../../scripts/lib/stable-relative-import-order.ts";

// @stll/folio-core publishes a source-mirrored dist: every source module maps
// 1:1 to a `dist/*.js` + `dist/*.d.ts` (`unbundle: true`). The package's
// `exports` expose a `"./*"` subpath wildcard onto that tree, so adapters
// (folio-react today; a Vue/Tauri adapter tomorrow) can import any core module
// by path — both in-repo from `./src/*` and, once published, from `./dist/*`.
//
// Tests are excluded from the build: they are not part of the published
// surface, and the React adapter resolves the few shared test utilities
// (`measure/__tests__/fakeTextMeasure`) from source in-repo, never from dist.
const entry = [
  "src/**/*.ts",
  "src/**/*.tsx",
  "!src/**/*.test.ts",
  "!src/**/*.test.tsx",
  "!src/**/__tests__/**",
];

const shared = {
  entry,
  // ESM-only: folio's own pervasive deps (better-result, marked) are ESM-only
  // packages with no `require` condition, so a CJS build would emit `require()`
  // calls that throw ERR_REQUIRE_ESM on consumers without require(ESM). ESM-only
  // also sidesteps the dual-package hazard entirely.
  format: ["esm"] as const,
  platform: "neutral" as const,
  outDir: "dist",
  // Source-mirrored output: one dist module per source module, inter-module
  // imports preserved as relative, every runtime dependency left external.
  unbundle: true,
};

// JS and declarations are emitted in two separate rolldown passes. Generating
// both in one build makes a type-only re-export module
// (`types/document.ts`, `export type * from "@stll/docx-core/model"`) collapse
// to an empty JS chunk that the declaration pass claims as a `.d.ts`; the JS
// then emits a runtime-helper import pointing at that `.d.ts`, which fails to
// load. Splitting the passes keeps each graph self-referential (JS imports
// `.js`, declarations import `.d.ts`).
//
// tsdown runs array configs concurrently, so neither pass may `clean` (it would
// race the other's output). The `build` script clears `dist` up front instead.
export default defineConfig([
  {
    ...shared,
    dts: false,
    clean: false,
    plugins: [stableRelativeImportOrder()],
  },
  { ...shared, dts: { emitDtsOnly: true }, clean: false },
]);
