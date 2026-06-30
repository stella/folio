import tailwindcss from "@tailwindcss/vite";
import react from "@vitejs/plugin-react";
import fs from "node:fs";
import path from "node:path";
import { defineConfig, type Plugin } from "vite";

const playgroundRoot = import.meta.dirname;
const repoRoot = path.resolve(playgroundRoot, "../..");
const fixturesDir = path.join(repoRoot, "tests/visual/fixtures");

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
  plugins: [tailwindcss(), react(), serveFixtures()],
  root: playgroundRoot,
  server: {
    port: 4200,
    strictPort: true,
    open: false,
  },
  build: {
    outDir: "dist",
  },
});
