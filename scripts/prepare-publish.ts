#!/usr/bin/env bun
// Transform a package.json from its in-repo "source" shape to the published
// "dist" shape, in place.
//
// The monorepo consumes these packages from source: `exports` point at
// `./src/*.ts`, so every consumer (Bun, tgo, Vite) resolves source directly
// with no aliases or conditions. The published package must instead expose the
// built artifacts — real `.d.ts` + `.js`, no source — so external consumers do
// not depend on our `.ts` or our bundler resolution.
//
// Run this after `bun run build`, immediately before `bun pm pack` /
// `bun publish`. It rewrites `exports` (deriving each dist target from the
// source path), `main`, `types`, and `files`. Restore the working tree
// afterward (`git checkout -- package.json`) — the publish workflow runs on an
// ephemeral checkout; the bootstrap script restores explicitly.
//
// Dist filenames are resolved against the freshly built `dist/` directory so
// this stays generic across both entry-naming schemes the packages use:
// path-preserving array entries (`./model` -> `./dist/model/document.js`) and
// flat object entries (`@stll/folio`'s `./core` -> `./dist/core.js`). Whichever
// file was actually emitted wins.

import { panic } from "better-result";
import { existsSync } from "node:fs";
import path from "node:path";

type JsEntry = { types: string; import: string };

const pkgDir = process.argv[2] ?? panic("usage: bun scripts/prepare-publish.ts <package-dir>");

const pkgPath = path.resolve(pkgDir, "package.json");
const pkg = await Bun.file(pkgPath).json();

// Drop the trailing extension so a base can be re-suffixed with the built one
// (`.js` / `.d.ts` / `.css`). `editor.css` -> `editor`, `model/document.ts` ->
// `model/document`, `core` -> `core`.
const stripExt = (file: string): string => file.slice(0, file.length - path.extname(file).length);

// Candidate dist bases for a source export, most-specific first:
//  - flat: the export subpath's own name (`./core` -> `core`, `.` -> `index`,
//    `./editor.css` -> `editor`)
//  - nested: the source tree mirrored (`./src/model/document.ts` ->
//    `model/document`)
// The first whose built file exists is used.
const distBases = (subpath: string, srcPath: string): string[] => {
  const flat = subpath === "." ? "index" : stripExt(subpath.replace(/^\.\//u, ""));
  const nested = stripExt(srcPath.replace(/^\.\/src\//u, ""));
  return flat === nested ? [flat] : [flat, nested];
};

// Resolve a built dist file from candidate bases + an extension, asserting it
// exists so a missing `bun run build` fails loudly rather than shipping a
// dangling export.
const resolveDist = (subpath: string, srcPath: string, ext: string): string => {
  for (const base of distBases(subpath, srcPath)) {
    const rel = `./dist/${base}${ext}`;
    if (existsSync(path.resolve(pkgDir, rel))) {
      return rel;
    }
  }
  return panic(
    `${pkg.name}: no built ${ext} for export "${subpath}" (run \`bun run build\` first)`,
  );
};

type ExportTarget = string | JsEntry;
const distExports: Record<string, ExportTarget> = {};
for (const [subpath, target] of Object.entries(pkg.exports)) {
  if (typeof target !== "string" || !target.startsWith("./src/")) {
    panic(
      `${pkg.name}: expected source export "${subpath}" to be a ./src/* string, got ${JSON.stringify(target)}`,
    );
  }

  // A subpath pattern (`"./*": "./src/*.ts"`) stands in for every built module:
  // map the `./src/` pattern to its `./dist/` JS + declaration counterparts
  // without a per-file existence check (the `*` cannot be globbed against one
  // built artifact). @stll/folio-core exposes its source-mirrored dist this way.
  if (subpath.includes("*")) {
    const base = target.replace(/^\.\/src\//u, "./dist/").replace(/\.[cm]?tsx?$/u, "");
    distExports[subpath] = { types: `${base}.d.ts`, import: `${base}.js` };
    continue;
  }

  // A CSS (or other asset) export publishes as a plain string to the bundled
  // dist asset — no `types`, no conditions.
  const ext = path.extname(target);
  if (ext === ".css") {
    distExports[subpath] = resolveDist(subpath, target, ".css");
    continue;
  }

  const jsBase = resolveDist(subpath, target, ".js").replace(/\.js$/u, "");
  distExports[subpath] = {
    types: `${jsBase}.d.ts`,
    import: `${jsBase}.js`,
  };
}

const root = distExports["."];
if (typeof root !== "object") {
  panic(`${pkg.name}: exports must include a JS "." entry`);
}

pkg.exports = distExports;
pkg.main = root.import;
pkg.types = root.types;
// "skills" ships TanStack Intent agent skills (skills/*/SKILL.md) for packages
// that have them; pack ignores entries that do not exist in a given package.
pkg.files = ["dist", "skills", "README.md", "LICENSE", "NOTICE.md"];

await Bun.write(pkgPath, `${JSON.stringify(pkg, null, 2)}\n`);
console.log(`prepared ${pkg.name}@${pkg.version} for publish (exports -> dist)`);
