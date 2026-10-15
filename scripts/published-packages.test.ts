import { expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { discoverPublishedPackages } from "./lib/published-packages";

test("new public workspace packages are discovered and built after their dependencies", () => {
  const root = mkdtempSync(path.join(tmpdir(), "folio-public-workspaces-"));
  try {
    writeFileSync(path.join(root, "package.json"), JSON.stringify({ workspaces: ["modules/*"] }));
    for (const [slug, manifest] of Object.entries({
      adapter: { name: "adapter", dependencies: { engine: "workspace:*" } },
      engine: { name: "engine" },
      playground: { name: "playground", private: true },
    })) {
      mkdirSync(path.join(root, "modules", slug), { recursive: true });
      writeFileSync(path.join(root, "modules", slug, "package.json"), JSON.stringify(manifest));
    }
    expect(discoverPublishedPackages(root).map(({ name }) => name)).toEqual(["engine", "adapter"]);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});
