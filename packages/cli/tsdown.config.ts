import { defineConfig } from "tsdown";

import { stableRelativeImportOrder } from "../../scripts/lib/stable-relative-import-order.ts";

// @stll/folio-cli publishes a source-mirrored dist like @stll/folio-agents:
// one `dist/*.js` + `dist/*.d.ts` per source module (`unbundle: true`), every
// runtime dependency left external. `src/bin.ts` keeps its `#!/usr/bin/env
// node` line, so `dist/bin.js` is the package's executable.
//
// Tests and their fixtures are excluded from the build.
const entry = ["src/**/*.ts", "!src/**/*.test.ts", "!src/**/__tests__/**"];

const shared = {
  entry,
  format: ["esm"] as const,
  // The CLI reads and writes files, spawns `git` for the default author, and
  // serves MCP over stdio: it targets Node, not a neutral runtime.
  platform: "node" as const,
  // `.js` like the other packages (the package is `"type": "module"`);
  // tsdown's Node default is `.mjs`, which prepare-publish would not find.
  fixedExtension: false,
  outDir: "dist",
  unbundle: true,
};

// JS and declarations are emitted in two separate passes, as in
// packages/agents/tsdown.config.ts; neither pass may `clean` because tsdown
// runs array configs concurrently, so the `build` script clears `dist`.
export default defineConfig([
  {
    ...shared,
    dts: false,
    clean: false,
    plugins: [stableRelativeImportOrder()],
  },
  { ...shared, dts: { emitDtsOnly: true }, clean: false },
]);
