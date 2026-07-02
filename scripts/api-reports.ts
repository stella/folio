#!/usr/bin/env bun
// Public-API surface snapshots for the published packages.
//
// For every published package (`@stll/folio-core`, `@stll/folio-react`) this
// walks the `exports` map, runs API Extractor over each subpath's built
// declaration file, and writes one normalized `<entry>.api.md` snapshot under
// `api-reports/<pkg>/`. The snapshots are committed so any change to the
// exported symbols or their signatures shows up as a reviewable diff.
//
// Usage:
//   bun scripts/api-reports.ts            # CI mode: fail on drift / missing dist
//   bun scripts/api-reports.ts --update   # write/refresh the committed snapshots
//   bun scripts/api-reports.ts --package core   # scope to one package
//
// Wired through package scripts as `api:check` (CI) and `api:update` (local).
// Requires a prior `bun run build` so the `dist/**/*.d.ts` inputs exist; CI
// reuses the build step's artifacts.
//
// Export -> dist mapping mirrors `scripts/prepare-publish.ts` (the canonical
// source->dist transform used at publish time): each `./src/<x>.ts` export
// resolves to the built `./dist/<base>.d.ts`, trying the flat subpath base
// first (`./messages` -> `dist/messages.d.ts`) then the mirrored source path
// (`./markdown` -> `dist/markdown/index.d.ts`), whichever was emitted.

import {
  CompilerState,
  Extractor,
  ExtractorConfig,
  ExtractorLogLevel,
} from "@microsoft/api-extractor";
import { existsSync, mkdirSync } from "node:fs";
import path from "node:path";

type PackageTarget = { slug: string; name: string; root: string };

const repoRoot = path.resolve(import.meta.dir, "..");

// The published packages, keyed by their `packages/<slug>` directory. Adding a
// new published package means adding one entry here.
const PACKAGES: PackageTarget[] = [
  { slug: "core", name: "@stll/folio-core", root: path.join(repoRoot, "packages/core") },
  { slug: "react", name: "@stll/folio-react", root: path.join(repoRoot, "packages/react") },
];

type Entry = { key: string; slug: string; dts: string };

// Drop the trailing extension so a base can be re-suffixed with `.d.ts`.
const stripExt = (file: string): string => file.slice(0, file.length - path.extname(file).length);

// Slug for a report filename: "." -> "index", "./prosemirror/attrs" ->
// "prosemirror-attrs".
const slugForKey = (key: string): string =>
  key === "." ? "index" : key.replace(/^\.\//u, "").replace(/\//gu, "-");

// Candidate dist declaration bases for a source export, most-specific first —
// the same order `prepare-publish.ts` uses to pick the emitted artifact.
const distBases = (subpath: string, srcPath: string): string[] => {
  const flat = subpath === "." ? "index" : stripExt(subpath.replace(/^\.\//u, ""));
  const nested = stripExt(srcPath.replace(/^\.\/src\//u, ""));
  return flat === nested ? [flat] : [flat, nested];
};

// Resolve the built `.d.ts` for a source export. Returns null when the export
// carries no type surface we can snapshot (see `entriesFor`).
const resolveDts = (packageRoot: string, subpath: string, srcPath: string): string | null => {
  for (const base of distBases(subpath, srcPath)) {
    const rel = `dist/${base}.d.ts`;
    if (existsSync(path.join(packageRoot, rel))) return rel;
  }
  return null;
};

// Turn a package's `exports` map into snapshot entries. Excluded, by design:
//   - `./package.json`            metadata, not a type surface
//   - subpath patterns (`./*`)    a wildcard passthrough to arbitrary internal
//                                 modules, not a curated public surface; there
//                                 is no single `.d.ts` to snapshot
//   - non-JS assets (`*.css`)     stylesheet, carries no declarations
const entriesFor = (pkg: PackageTarget): { entries: Entry[]; missing: string[] } => {
  const pkgJson = require(path.join(pkg.root, "package.json")) as {
    exports: Record<string, unknown>;
  };
  const entries: Entry[] = [];
  const missing: string[] = [];
  for (const [key, srcPath] of Object.entries(pkgJson.exports)) {
    if (key === "./package.json" || key.includes("*")) continue;
    // In-repo exports are plain `./src/*.ts` strings; `prepare-publish.ts`
    // panics on any other shape at publish time. Fail loudly here too, so a
    // future refactor to conditional-exports objects breaks this gate visibly
    // instead of silently dropping entries from the snapshot.
    if (typeof srcPath !== "string") {
      console.error(
        `${pkg.name}: export "${key}" is not a plain source-path string ` +
          `(got ${JSON.stringify(srcPath)}).\n` +
          `In-repo exports must point at ./src/*.ts (see scripts/prepare-publish.ts); ` +
          `teach scripts/api-reports.ts about the new shape before changing this.`,
      );
      process.exit(1);
    }
    if (path.extname(srcPath) === ".css") continue;
    const dts = resolveDts(pkg.root, key, srcPath);
    if (dts) {
      entries.push({ key, slug: slugForKey(key), dts });
    } else {
      missing.push(`${key} (${srcPath})`);
    }
  }
  return { entries, missing };
};

const reportDirFor = (pkg: PackageTarget): string => path.join(repoRoot, "api-reports", pkg.slug);
const tempDirFor = (pkg: PackageTarget): string =>
  path.join(repoRoot, ".cache", "api-reports", pkg.slug);

type BuildConfigOptions = { pkg: PackageTarget; entry: Entry; reportDir: string; tempDir: string };

const buildConfig = ({ pkg, entry, reportDir, tempDir }: BuildConfigOptions): ExtractorConfig => {
  const packageJsonFullPath = path.join(pkg.root, "package.json");
  return ExtractorConfig.prepare({
    configObject: {
      mainEntryPointFilePath: path.join(pkg.root, entry.dts),
      apiReport: {
        enabled: true,
        reportFolder: reportDir,
        reportFileName: `${entry.slug}.api.md`,
        reportTempFolder: tempDir,
      },
      docModel: { enabled: false },
      dtsRollup: { enabled: false },
      tsdocMetadata: { enabled: false },
      compiler: { tsconfigFilePath: path.join(pkg.root, "tsconfig.json") },
      messages: {
        // Unbundled dist emits one `.d.ts` per source module, so barrels
        // re-export types that live in sibling files: API Extractor still
        // follows and inlines them, but flags them "forgotten." Silence the
        // noise. `ae-missing-release-tag` is a warning, not an error, because
        // folio does not annotate `@public`/`@internal` on every symbol; the
        // snapshot itself is the contract, not the release tags.
        extractorMessageReporting: {
          "ae-forgotten-export": { logLevel: "none" },
          "ae-missing-release-tag": { logLevel: "warning" },
        },
        tsdocMessageReporting: { "tsdoc-undefined-tag": { logLevel: "none" } },
      },
      projectFolder: pkg.root,
    },
    configObjectFullPath: packageJsonFullPath,
    packageJsonFullPath,
  });
};

type RunResult = { errors: number; drifted: Entry[] };

const runPackage = (pkg: PackageTarget, isLocal: boolean): RunResult => {
  const { entries, missing } = entriesFor(pkg);
  const reportDir = reportDirFor(pkg);
  const tempDir = tempDirFor(pkg);
  mkdirSync(reportDir, { recursive: true });
  mkdirSync(tempDir, { recursive: true });

  if (missing.length > 0) {
    console.error(`\nMissing built .d.ts for ${missing.length} export(s) in ${pkg.name}:`);
    for (const m of missing) console.error(`  - ${m}`);
    console.error(`\nFix: bun run build:${pkg.slug}`);
    process.exit(1);
  }
  if (entries.length === 0) {
    console.error(`No snapshot entries for ${pkg.name}. Run \`bun run build:${pkg.slug}\` first.`);
    process.exit(1);
  }

  // Share one CompilerState across every entry so the tsconfig is parsed and
  // the dist tree walked once, not once per subpath.
  const firstConfig = buildConfig({ pkg, entry: entries[0], reportDir, tempDir });
  const compilerState = CompilerState.create(firstConfig, {
    additionalEntryPoints: entries.slice(1).map((e) => path.join(pkg.root, e.dts)),
  });

  let errors = 0;
  const drifted: Entry[] = [];
  for (const entry of entries) {
    const result = Extractor.invoke(buildConfig({ pkg, entry, reportDir, tempDir }), {
      localBuild: isLocal,
      showVerboseMessages: false,
      compilerState,
      // Suppress console/info chatter (bundled-TS version banners, per-entry
      // "copy the file" drift hints — the driver prints its own drift summary)
      // and warning spam (`ae-missing-release-tag` on every untagged symbol).
      // Error-severity messages fall through to API Extractor's default
      // handler so their text is printed. Note `errorCount`/`warningCount`
      // are incremented before the `handled` check (MessageRouter), so even a
      // handled error still fails the run via `totalErrors` below; letting
      // errors through only restores the diagnostic text.
      messageCallback: (message) => {
        if (message.logLevel !== ExtractorLogLevel.Error) {
          message.handled = true;
        }
      },
    });
    errors += result.errorCount;
    if (!isLocal && result.apiReportChanged) drifted.push(entry);
  }

  console.log(`${pkg.name}: ${entries.length} entries, ${errors} errors`);
  return { errors, drifted };
};

const args = process.argv.slice(2);
const isLocal = args.includes("--update") || args.includes("--local");
const pkgArgIdx = args.indexOf("--package");
const pkgArg = pkgArgIdx !== -1 ? args[pkgArgIdx + 1] : null;

const targets = pkgArg ? PACKAGES.filter((p) => p.slug === pkgArg || p.name === pkgArg) : PACKAGES;
if (pkgArg && targets.length === 0) {
  console.error(`Unknown package '${pkgArg}'. Known: ${PACKAGES.map((p) => p.slug).join(", ")}`);
  process.exit(1);
}

let totalErrors = 0;
const allDrifted: Entry[] = [];
for (const pkg of targets) {
  const { errors, drifted } = runPackage(pkg, isLocal);
  totalErrors += errors;
  allDrifted.push(...drifted);
}

if (allDrifted.length > 0) {
  console.error(`\nPublic-API surface drift in ${allDrifted.length} entr(y/ies):`);
  for (const e of allDrifted) console.error(`  - ${e.slug} (${e.key})`);
  console.error(`\nThe exported API changed but the committed snapshot did not.`);
  console.error(`Run \`bun run api:update\`, review the diff under api-reports/, and commit it.`);
  process.exit(1);
}
if (totalErrors > 0) process.exit(1);

if (isLocal) console.log("\nSnapshots written to api-reports/. Review and commit the diff.");
else console.log("\nPublic-API snapshots up to date.");
