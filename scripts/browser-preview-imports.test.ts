import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

// Scan all browser harness/host sources, including future specs and helpers.
// Source imports must be resolved by the bundler, never by a dev server URL.
const DEV_ONLY_URL =
  /\/@(?:fs|id|vite|react-refresh)(?:\/|\b)|\/node_modules\/\.vite\/|["'`]\/(?:src|packages|tests)\/|["'`]\/[^"'`\s]+\.(?:ts|tsx|vue)(?:[?"'`])/u;
const roots = ["tests/visual", "packages/playground", "packages/playground-vue"];
const repo = resolve(import.meta.dir, "..");

test("browser harness and playground sources contain no dev-only transport URLs", () => {
  const violations: string[] = [];
  for (const root of roots) {
    for (const file of new Bun.Glob("**/*.{ts,tsx,vue,html}").scanSync({
      cwd: resolve(repo, root),
      onlyFiles: true,
    })) {
      if (file.startsWith("node_modules/") || file.startsWith("dist/")) continue;
      const source = readFileSync(resolve(repo, root, file), "utf8");
      if (DEV_ONLY_URL.test(source)) violations.push(`${root}/${file}`);
    }
  }
  expect(violations).toEqual([]);
});

test("guard covers filesystem, transformed imports, Vite client and optimized dependencies", () => {
  for (const prefix of ["fs", "id", "vite"]) {
    expect(DEV_ONLY_URL.test(`import("/@${prefix}/helper.ts")`)).toBe(true);
    expect(DEV_ONLY_URL.test(`const url = \`/@${prefix}\${absolutePath}\``)).toBe(true);
  }
  expect(DEV_ONLY_URL.test('import("/node_modules/.vite/deps/helper.js")')).toBe(true);
  for (const url of [
    "/src/helper.js",
    "/packages/core/helper.js",
    "/tests/helper.js",
    "/helper.ts",
    "/helper.tsx?import",
    "/helper.vue",
    "/@react-refresh",
  ]) {
    expect(DEV_ONLY_URL.test(`import("${url}")`)).toBe(true);
  }
  expect(DEV_ONLY_URL.test('import("./helper")')).toBe(false);
  expect(DEV_ONLY_URL.test('fetch("/assets/font.woff")')).toBe(false);
});
