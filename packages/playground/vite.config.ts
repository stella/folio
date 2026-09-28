import tailwindcss from "@tailwindcss/vite";
import react from "@vitejs/plugin-react";
import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import {
  defineConfig,
  searchForWorkspaceRoot,
  type Plugin,
  type PreviewServer,
  type ViteDevServer,
} from "vite";

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

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const FIXTURE_PREFIX = "/fixtures/";

// Serve the repo's visual-test fixtures at `/fixtures/<name>.docx`. The editor
// fetches a fixture from this path when the page is opened with `?file=<name>`,
// and both the rendering and interaction specs depend on it. A dev-server
// middleware (rather than a copied/symlinked `public/` dir) keeps
// `tests/visual/fixtures` the single source of truth.
function serveFixtures(): Plugin {
  return {
    name: "folio-serve-fixtures",
    configureServer(server) {
      installFixtureMiddleware(server, "no-cache");
    },
    configurePreviewServer(server) {
      installFixtureMiddleware(server, "public, max-age=3600");
    },
  };
}

function installFixtureMiddleware(
  server: PreviewServer | ViteDevServer,
  cacheControl: string,
): void {
  server.middlewares.use((req, res, next) => {
    if (!req.url || !req.url.startsWith(FIXTURE_PREFIX)) {
      next();
      return;
    }
    const requested = req.url.slice(FIXTURE_PREFIX.length).split("?")[0] ?? "";
    let name: string;
    try {
      name = decodeURIComponent(requested);
    } catch {
      res.statusCode = 400;
      res.end("Invalid fixture path");
      return;
    }
    if (
      !name ||
      name !== path.basename(name) ||
      name.includes("/") ||
      name.includes("\\") ||
      name.includes("..")
    ) {
      res.statusCode = 400;
      res.end("Invalid fixture path");
      return;
    }
    fs.readFile(path.join(fixturesDir, name), (error, data) => {
      if (error) {
        res.statusCode = 404;
        res.end(`Fixture not found: ${name}`);
        return;
      }
      res.setHeader("Content-Type", DOCX_MIME);
      res.setHeader("Cache-Control", cacheControl);
      res.end(data);
    });
  });
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
  build: {
    outDir: "dist",
  },
});
