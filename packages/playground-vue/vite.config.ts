import vue from "@vitejs/plugin-vue";
import path from "node:path";
import { defineConfig, type Plugin, searchForWorkspaceRoot } from "vite";

import { fixtureMiddleware } from "../../scripts/playground-fixtures";

const playgroundRoot = import.meta.dirname;
const repoRoot = path.resolve(playgroundRoot, "../..");
const fixturesDir = path.join(repoRoot, "tests/visual/fixtures");

// Both lifecycle hooks serve the same repository fixture bytes and path policy.
function serveFixtures(): Plugin {
  return {
    name: "folio-serve-fixtures",
    configureServer(server) {
      server.middlewares.use(
        fixtureMiddleware({ fixturesDir, cacheControl: "no-cache" }),
      );
    },
    configurePreviewServer(server) {
      server.middlewares.use(
        fixtureMiddleware({
          fixturesDir,
          cacheControl: "public, max-age=3600",
        }),
      );
    },
  };
}

export default defineConfig({
  plugins: [vue(), serveFixtures()],
  root: playgroundRoot,
  resolve: {
    // Serve workspace source, including the font entry, without a stale dist build.
    alias: [
      {
        find: "@stll/folio-vue/editor.css",
        replacement: path.resolve(
          repoRoot,
          "packages/vue/src/styles/playground.css",
        ),
      },
      {
        find: /^@stll\/folio-vue$/u,
        replacement: path.resolve(repoRoot, "packages/vue/src/index.ts"),
      },
    ],
  },
  // Mirror the React playground: when launched by the parity harness, serve
  // workspace packages as live source instead of a cached pre-bundle.
  ...(process.env["FOLIO_PLAYGROUND_PORT"]
    ? { optimizeDeps: { exclude: ["@stll/folio-core", "@stll/folio-vue"] } }
    : {}),
  server: {
    // Distinct from the React playground's 4200 so both can run in parallel for
    // the cross-adapter parity project.
    port: Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4201,
    strictPort: true,
    open: false,
    fs: {
      // The bundled faces come from packages/fonts, inside the workspace.
      allow: [searchForWorkspaceRoot(playgroundRoot)],
    },
  },
  preview: {
    port: Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4201,
    strictPort: true,
    open: false,
  },
  build: {
    outDir: "dist",
  },
});
