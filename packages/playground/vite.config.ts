import tailwindcss from "@tailwindcss/vite";
import react from "@vitejs/plugin-react";
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
  plugins: [tailwindcss(), react(), serveFixtures()],
  root: playgroundRoot,
  // When launched by the parity harness (which sets FOLIO_PLAYGROUND_PORT),
  // serve the workspace packages as live source instead of pre-bundling them
  // into `.vite/deps`. Vite's dep-optimizer caches the bundled snapshot on
  // disk and does NOT re-bundle when workspace source changes, so a fresh
  // server would otherwise still serve stale `@stll/folio-core` — silently
  // making the parity feedback loop measure old layout code. Excluding keeps
  // every parity run current. Normal manual dev keeps pre-bundling for speed.
  ...(process.env["FOLIO_PLAYGROUND_PORT"]
    ? { optimizeDeps: { exclude: ["@stll/folio-core", "@stll/folio-react"] } }
    : {}),
  server: {
    // Default 4200 for manual dev; the parity harness overrides this per
    // worktree via FOLIO_PLAYGROUND_PORT so parallel worktrees don't collide.
    port: Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200,
    strictPort: true,
    open: false,
    fs: {
      // The bundled faces come from packages/fonts, inside the workspace.
      allow: [searchForWorkspaceRoot(playgroundRoot)],
    },
  },
  preview: {
    port: Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200,
    strictPort: true,
    open: false,
  },
  build: {
    outDir: "dist",
  },
});
