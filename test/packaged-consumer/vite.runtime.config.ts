import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

// Production SPA build of the runtime smoke consumer (`index.html` -> `app.tsx`).
// Unlike the lib build in `vite.config.ts` (bundler-fidelity only, not run), this
// build is SERVED and driven in a browser by `smoke.spec.ts`.
//
// A default SPA build bundles every imported dependency, so the packed
// @stll/folio-* tarballs — and the font-metrics worker inside folio-core — all
// go through the bundler here. Vite's `vite:worker-import-meta-url` plugin emits
// the worker as its own chunk (`assets/font-metrics.worker-*.js`); the served
// build then constructs it at runtime, which is exactly what the smoke exercises.
export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "dist-runtime",
    // A module worker needs a browser target that supports `import` in workers.
    target: "es2022",
  },
  preview: {
    port: 4300,
    strictPort: true,
  },
});
