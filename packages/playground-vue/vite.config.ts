import vue from "@vitejs/plugin-vue";
import fs from "node:fs";
import path from "node:path";
import { defineConfig, type Plugin } from "vite";

const playgroundRoot = import.meta.dirname;
const repoRoot = path.resolve(playgroundRoot, "../..");
const fixturesDir = path.join(repoRoot, "tests/visual/fixtures");

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const FIXTURE_PREFIX = "/fixtures/";

// Serve the repo's visual-test fixtures at `/fixtures/<name>.docx` — the same
// middleware the React playground uses, so both adapters load identical bytes
// with `?file=<name>`. `tests/visual/fixtures` stays the single source of truth.
function serveFixtures(): Plugin {
  return {
    name: "folio-serve-fixtures",
    configureServer(server) {
      server.middlewares.use((req, res, next) => {
        if (!req.url || !req.url.startsWith(FIXTURE_PREFIX)) {
          next();
          return;
        }
        const requested = req.url.slice(FIXTURE_PREFIX.length).split("?")[0] ?? "";
        const name = decodeURIComponent(requested);
        if (!name || name.includes("/") || name.includes("..")) {
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
          res.setHeader("Cache-Control", "no-cache");
          res.end(data);
        });
      });
    },
  };
}

export default defineConfig({
  plugins: [vue(), serveFixtures()],
  root: playgroundRoot,
  resolve: {
    alias: {
      // The published `@stll/folio-vue` `.` export points at `dist`, which is a
      // build artifact. Serve the workspace source directly so the playground
      // (and the parity e2e loop) never measures a stale build.
      "@stll/folio-vue": path.resolve(repoRoot, "packages/vue/src/index.ts"),
    },
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
  },
  build: {
    outDir: "dist",
  },
});
