import tailwindcss from "@tailwindcss/vite";
import react from "@vitejs/plugin-react";
import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import { fixtureMiddleware } from "../../scripts/playground-fixtures";
import { defineConfig, searchForWorkspaceRoot, type Plugin } from "vite";

const playgroundRoot = import.meta.dirname;
const repoRoot = path.resolve(playgroundRoot, "../..");
const fixturesDir = path.join(repoRoot, "tests/visual/fixtures");

const reactPackageJsonPath = path.join(repoRoot, "packages/react/package.json");

/**
 * The real directories of the fontsource packages `@stll/folio-react` bundles.
 *
 * Bun's isolated global store keeps package files outside the checkout, and
 * the dev server refuses to serve anything outside the workspace, so every
 * bundled face failed to load (status `error`) and the playground measured and
 * painted in system fallbacks. Derived from the adapter's dependencies so a
 * new bundled face is served without editing this list.
 */
function bundledFontPackageDirs(): string[] {
  const reactPackage: { dependencies?: Record<string, string> } = JSON.parse(
    fs.readFileSync(reactPackageJsonPath, "utf8"),
  );
  const requireFromReact = createRequire(reactPackageJsonPath);
  return Object.keys(reactPackage.dependencies ?? {})
    .filter((name) => name.startsWith("@fontsource/"))
    .map((name) => path.dirname(fs.realpathSync(requireFromReact.resolve(`${name}/package.json`))));
}

// Both lifecycle hooks serve the same repository fixture bytes and path policy.
function serveFixtures(): Plugin {
  return {
    name: "folio-serve-fixtures",
    configureServer(server) {
      server.middlewares.use(fixtureMiddleware({ fixturesDir, cacheControl: "no-cache" }));
    },
    configurePreviewServer(server) {
      server.middlewares.use(
        fixtureMiddleware({ fixturesDir, cacheControl: "public, max-age=3600" }),
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
      allow: [searchForWorkspaceRoot(playgroundRoot), ...bundledFontPackageDirs()],
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
