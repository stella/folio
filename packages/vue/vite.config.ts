import vue from "@vitejs/plugin-vue";
import { resolve } from "node:path";
import { defineConfig } from "vite";
import dts from "vite-plugin-dts";

// Library build for @stll/folio-vue. Vite (not tsdown) because the package
// ships .vue SFCs that need the @vitejs/plugin-vue compiler step — the rest of
// the monorepo builds plain .ts with tsdown. External: vue, prosemirror-*, and
// @stll/folio-core — consumers bring those.
export default defineConfig({
  plugins: [
    vue(),
    dts({
      include: ["src/**/*"],
      // Keep type-level conformance tests out of the published tarball.
      exclude: ["src/**/__tests__/**", "src/**/*.test-d.ts", "src/**/*.test.ts"],
      // Pin the entry root so multi-entry builds flatten declarations to
      // dist/index.d.ts + dist/ui.d.ts (auto-detect drifts to a parent dir
      // once core's workspace types enter the graph).
      entryRoot: "src",
      // Keep the workspace package name (`@stll/folio-core`) intact in
      // published declarations instead of rewriting to `../../core/src/*`.
      pathsToAliases: false,
      // `.d.ts.map` points at source .ts files absent from the tarball.
      compilerOptions: { declarationMap: false },
    }),
  ],
  build: {
    lib: {
      // Keep public subpaths backed by real JS chunks so consumers can import
      // composables/plugin APIs without dragging in the full editor shell.
      entry: {
        index: resolve(__dirname, "src/index.ts"),
        ui: resolve(__dirname, "src/ui.ts"),
        composables: resolve(__dirname, "src/composables/index.ts"),
        dialogs: resolve(__dirname, "src/components/dialogs/index.ts"),
        styles: resolve(__dirname, "src/styles/index.ts"),
        messages: resolve(__dirname, "src/i18n/messages.ts"),
      },
      formats: ["es", "cjs"],
      fileName: (format, entryName) => `${entryName}.${format === "es" ? "js" : "cjs"}`,
      cssFileName: "folio-vue",
    },
    rollupOptions: {
      external: ["vue", /^@stll\/folio-core(\/.*)?$/, /^prosemirror-/],
    },
    emptyOutDir: true,
  },
});
