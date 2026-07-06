import { defineConfig } from "tsdown";

// @stll/folio-agents publishes a source-mirrored dist, same shape as
// @stll/folio-core: every source module maps 1:1 to a `dist/*.js` +
// `dist/*.d.ts` (`unbundle: true`), and the package's `exports` expose a
// `"./*"` subpath wildcard onto that tree.
//
// Tests are excluded from the build: they are not part of the published
// surface.
const entry = ["src/**/*.ts", "!src/**/*.test.ts", "!src/**/__tests__/**"];

const shared = {
  entry,
  // ESM-only: mirrors @stll/folio-core (this package's only dependency is
  // ESM-only itself), sidestepping the dual-package hazard entirely.
  format: ["esm"] as const,
  platform: "neutral" as const,
  outDir: "dist",
  // Source-mirrored output: one dist module per source module, inter-module
  // imports preserved as relative, every runtime dependency (@stll/folio-core)
  // left external.
  unbundle: true,
};

// JS and declarations are emitted in two separate rolldown passes — see
// packages/core/tsdown.config.ts for why (a type-only re-export module can
// otherwise collapse into a JS chunk the declaration pass misclaims as a
// `.d.ts`). Neither pass may `clean` (tsdown runs array configs
// concurrently); the `build` script clears `dist` up front instead.
export default defineConfig([
  { ...shared, dts: false, clean: false },
  { ...shared, dts: { emitDtsOnly: true }, clean: false },
]);
