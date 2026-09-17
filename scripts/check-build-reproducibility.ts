#!/usr/bin/env bun
// Reproducible-build gate: a published package must emit byte-identical
// JavaScript *and* declarations when its release build runs twice.
//
// Usage:
//   bun scripts/check-build-reproducibility.ts
//   bun scripts/check-build-reproducibility.ts --packages core,docx-core
//
// Wired through package scripts as `check:build-reproducibility`.
//
// Each package is built by spawning `bun run build` in its own directory — the
// entry point `prepack` uses, so what the gate compares is the artifact the
// release publishes. Driving tsdown's programmatic API instead would only ever
// reach the JavaScript pass: tsdown re-reads `tsdown.config.ts` from the
// working directory and re-applies that pass's plugin to declaration chunks,
// which rolldown then fails to parse.
//
// No build here accepts an out-dir override — core/agents/docx-core pin
// `outDir` in tsdown.config.ts, react appends a CSS step that writes into
// `dist`, vue builds through Vite and nuxt through nuxt-module-build — so the
// first build's `dist` is moved aside, the second build recreates it, and the
// two trees are compared. The second build's `dist` is left in place: later CI
// steps (`api:check`, `validate-dist`) read it.

import { existsSync } from "node:fs";
import { mkdtemp, rename, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import {
  firstBuildTreeMismatch,
  renderBuildTreeMismatch,
  type BuildTree,
} from "./lib/build-tree-diff";
import { PUBLISHED_PACKAGES, type PublishedPackage } from "./lib/published-packages";

const MAX_DIFF_LINES = 200;

// Declarations and JavaScript across every module-specific extension the six
// build tools emit: nuxt-module-build writes `.mjs`/`.d.mts`, Vite adds `.cjs`.
// Matching only `.js`/`.d.ts` would make the nuxt run vacuous.
const COMPARED_EXTENSIONS = [".d.ts", ".d.mts", ".d.cts", ".js", ".mjs", ".cjs"] as const;

const isCompared = (file: string): boolean =>
  COMPARED_EXTENSIONS.some((extension) => file.endsWith(extension));

const readTree = async (directory: string): Promise<BuildTree> => {
  const files = Array.from(new Bun.Glob("**/*").scanSync({ cwd: directory, onlyFiles: true }))
    .filter(isCompared)
    .toSorted();
  const tree = new Map<string, string>();
  for (const file of files) {
    tree.set(file, await Bun.file(path.join(directory, file)).text());
  }
  return tree;
};

type BuildFailure = { exitCode: number; output: string };

/** A build that wrote to neither stream leaves its sink file uncreated. */
const readLog = async (file: string): Promise<string> =>
  existsSync(file) ? await Bun.file(file).text() : "";

const runBuild = async (
  pkg: PublishedPackage,
  logDirectory: string,
  attempt: number,
): Promise<BuildFailure | null> => {
  const stdoutPath = path.join(logDirectory, `${pkg.slug}-${attempt}-stdout.log`);
  const stderrPath = path.join(logDirectory, `${pkg.slug}-${attempt}-stderr.log`);
  // Piped stdout truncates around 768 KiB in Bun.spawnSync; a file sink does
  // not, and a truncated build log is worthless when the build is what failed.
  const { exitCode } = Bun.spawnSync(["bun", "run", "build"], {
    cwd: pkg.root,
    stdout: Bun.file(stdoutPath),
    stderr: Bun.file(stderrPath),
  });
  if (exitCode === 0) return null;
  return { exitCode, output: `${await readLog(stdoutPath)}${await readLog(stderrPath)}` };
};

const reportBuildFailure = (pkg: PublishedPackage, failure: BuildFailure): void => {
  console.error(`\n${pkg.name}: \`bun run build\` exited ${failure.exitCode}.`);
  console.error(failure.output);
};

/** True when the package's two builds agreed. */
const checkPackage = async (pkg: PublishedPackage, tempRoot: string): Promise<boolean> => {
  const distDirectory = path.join(pkg.root, "dist");
  const baselineDirectory = path.join(tempRoot, `${pkg.slug}-baseline`);
  const startedNs = Bun.nanoseconds();

  const firstFailure = await runBuild(pkg, tempRoot, 1);
  if (firstFailure) {
    reportBuildFailure(pkg, firstFailure);
    return false;
  }
  if (!existsSync(distDirectory)) {
    console.error(`\n${pkg.name}: \`bun run build\` emitted no dist/.`);
    return false;
  }
  await rename(distDirectory, baselineDirectory);

  const secondFailure = await runBuild(pkg, tempRoot, 2);
  if (secondFailure) {
    // Put a dist back before leaving: a half-run gate must not strip the
    // working tree of artifacts the next step reads.
    await rm(distDirectory, { recursive: true, force: true });
    await rename(baselineDirectory, distDirectory);
    reportBuildFailure(pkg, secondFailure);
    return false;
  }

  const baseline = await readTree(baselineDirectory);
  const candidate = await readTree(distDirectory);
  if (baseline.size === 0) {
    console.error(
      `\n${pkg.name}: dist/ holds no declaration or JavaScript files, so the gate ` +
        `checked nothing. Teach COMPARED_EXTENSIONS about what this build emits.`,
    );
    return false;
  }

  const mismatch = firstBuildTreeMismatch({ baseline, candidate });
  if (mismatch) {
    console.error(`\n${pkg.name} is not reproducible; first file that differs:\n`);
    console.error(renderBuildTreeMismatch({ mismatch, maxDiffLines: MAX_DIFF_LINES }));
    return false;
  }

  const seconds = ((Bun.nanoseconds() - startedNs) / 1e9).toFixed(1);
  console.log(`${pkg.name}: ${baseline.size} files identical across two builds (${seconds}s)`);
  return true;
};

const args = process.argv.slice(2);
const packagesFlag = args.indexOf("--packages");
const requested = packagesFlag === -1 ? null : args.at(packagesFlag + 1);
if (packagesFlag !== -1 && !requested) {
  console.error("--packages requires a comma-separated list of package slugs.");
  process.exit(1);
}

const known = PUBLISHED_PACKAGES.map(({ slug }) => slug).join(", ");
const requestedSlugs = requested
  ? requested
      .split(",")
      .map((slug) => slug.trim())
      .filter((slug) => slug.length > 0)
  : null;
if (requestedSlugs?.length === 0) {
  console.error(`--packages named no package. Known: ${known}`);
  process.exit(1);
}
const unknown = (requestedSlugs ?? []).filter(
  (slug) => !PUBLISHED_PACKAGES.some((pkg) => pkg.slug === slug),
);
if (unknown.length > 0) {
  console.error(`Unknown package(s) ${unknown.join(", ")}. Known: ${known}`);
  process.exit(1);
}

// Kept in release-build order rather than the order the flag listed them: a
// later package's build reads an earlier one's dist.
const targets = requestedSlugs
  ? PUBLISHED_PACKAGES.filter(({ slug }) => requestedSlugs.includes(slug))
  : PUBLISHED_PACKAGES;

const tempRoot = await mkdtemp(path.join(tmpdir(), "folio-reproducible-build-"));
const startedNs = Bun.nanoseconds();
let failed = 0;
try {
  for (const pkg of targets) {
    if (!(await checkPackage(pkg, tempRoot))) failed += 1;
  }
} finally {
  await rm(tempRoot, { recursive: true, force: true });
}

const totalSeconds = ((Bun.nanoseconds() - startedNs) / 1e9).toFixed(1);
if (failed > 0) {
  console.error(
    `\n${failed} of ${targets.length} package(s) failed the reproducible-build gate ` +
      `(${totalSeconds}s).`,
  );
  process.exit(1);
}
console.log(`\n${targets.length} package(s) build reproducibly (${totalSeconds}s).`);
