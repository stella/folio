#!/usr/bin/env bun
/**
 * Run the property tests of the areas a change touches at a raised run count.
 *
 * PR CI runs every property at its own `numRuns` (factor 1) inside the main
 * test step. A property that fails on a few percent of seeds can pass there and
 * redden unrelated PRs later, so this step reruns the property files of the
 * areas the branch touched with `PROPERTY_TEST_NUM_RUNS_FACTOR` raised (5 by
 * default): a change to `packages/core/src/compare/**` searches the compare
 * properties five times as hard, and nothing else pays for it.
 *
 * An area is `<package>/<first directory under src>` (`core/ai-edits`,
 * `core/compare`, `docx-core/markdown`, ...), plus the couplings below where
 * one directory's code is exercised by another's properties. Changed source
 * selects areas; a changed property file runs itself, and a change to
 * test/property-seeds.json runs the files it names.
 *
 * Usage:
 *   bun scripts/property-areas.ts [--base <ref>] [--factor <n>] [--dry-run]
 *   bun scripts/property-areas.ts --all            (every property file)
 *
 * `--base` (default `origin/main`) is diffed from its merge base with HEAD.
 * `--factor` defaults to `PROPERTY_TEST_NUM_RUNS_FACTOR`, else 5 for a change
 * and 1 for `--all` (the nightly sets the variable). `--dry-run` prints the
 * selection without running it.
 */

import { $ } from "bun";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const PACKAGES_DIR = path.join(REPO_ROOT, "packages");
const ROOT_TEST_DIRS = ["scripts", "test"] as const;
const PROPERTY_FILE = /\.test\.tsx?$/;
const PROPERTY_SEEDS_FILE = "test/property-seeds.json";

/**
 * Changed-path prefixes whose code another area's properties exercise, on top
 * of the area the path itself belongs to.
 */
export const COUPLINGS: readonly { prefix: string; areas: readonly string[] }[] = [
  // The comparison's internals live under internal/compare.
  { prefix: "packages/core/src/internal/compare/", areas: ["core/compare"] },
  // Whole-story revision resolution backs accept/reject in the reviewer and
  // the editor: the tracked-resolution and headless-review properties drive it.
  { prefix: "packages/core/src/internal/", areas: ["core/ai-edits", "core/prosemirror"] },
  // Tracked changes are ProseMirror marks resolved by the same commands.
  { prefix: "packages/core/src/prosemirror/", areas: ["core/ai-edits"] },
  // Both markdown readers share the docx-core compiler.
  { prefix: "packages/docx-core/src/markdown/", areas: ["core/markdown"] },
  { prefix: "packages/core/src/markdown/", areas: ["docx-core/markdown"] },
  // The comparison saves and reopens through the docx reader and writer.
  { prefix: "packages/core/src/docx/serializer/", areas: ["core/compare"] },
];

type PropertyFile = { area: string; packageDir: string; file: string };

const walk = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    if (entry.name === "node_modules" || entry.name.startsWith(".")) return [];
    const absolute = path.join(dir, entry.name);
    return entry.isDirectory() ? walk(absolute) : [absolute];
  });

const drivesProperty = (file: string): boolean => {
  const source = ts.createSourceFile(
    file,
    readFileSync(file, "utf8"),
    ts.ScriptTarget.Latest,
    true,
  );
  let found = false;
  const visit = (node: ts.Node): void => {
    if (ts.isCallExpression(node)) {
      const callee = node.expression;
      found ||=
        (ts.isIdentifier(callee) && callee.text === "assertProperty") ||
        (ts.isPropertyAccessExpression(callee) &&
          ts.isIdentifier(callee.expression) &&
          callee.expression.text === "fc" &&
          (callee.name.text === "assert" || callee.name.text === "check"));
    }
    if (!found) ts.forEachChild(node, visit);
  };
  visit(source);
  return found;
};

/** `packages/<pkg>/src/<segment>/...` → `<pkg>/<segment>`; a file directly under src → `<pkg>`. */
export const areaOf = (repoPath: string): string | undefined => {
  const match = /^packages\/([^/]+)\/src\/(?:([^/]+)\/)?/.exec(repoPath);
  if (match === null) return undefined;
  return match[2] === undefined ? match[1] : `${match[1]}/${match[2]}`;
};

/** Every package and root test file that drives a fast-check property. */
export const propertyFiles = (): PropertyFile[] => {
  const packages = readdirSync(PACKAGES_DIR).flatMap((pkg) => {
    const src = path.join(PACKAGES_DIR, pkg, "src");
    let files: string[];
    try {
      files = walk(src);
    } catch {
      return [];
    }
    return files
      .filter((file) => PROPERTY_FILE.test(file) && drivesProperty(file))
      .map((file) => {
        const repoPath = path.relative(REPO_ROOT, file).replaceAll("\\", "/");
        return {
          area: areaOf(repoPath) as string,
          packageDir: `packages/${pkg}`,
          file: repoPath,
        };
      })
      .toSorted((a, b) => a.file.localeCompare(b.file));
  });
  const rootTests = ROOT_TEST_DIRS.flatMap((root) =>
    walk(path.join(REPO_ROOT, root))
      .filter((file) => PROPERTY_FILE.test(file) && drivesProperty(file))
      .map((file) => ({
        area: root,
        packageDir: ".",
        file: path.relative(REPO_ROOT, file).replaceAll("\\", "/"),
      })),
  );
  return [...packages, ...rootTests].toSorted((a, b) => a.file.localeCompare(b.file));
};

/** The areas a set of changed repo paths touches. */
export const touchedAreas = (changed: readonly string[]): Set<string> => {
  const areas = new Set<string>();
  for (const file of changed) {
    // A test file is not code under test: a changed property file reruns
    // itself (selectPropertyFiles), a changed example test nothing more.
    if (PROPERTY_FILE.test(file)) continue;
    const own = areaOf(file);
    if (own !== undefined) areas.add(own);
    for (const root of ROOT_TEST_DIRS) {
      if (file.startsWith(`${root}/`)) areas.add(root);
    }
    for (const { prefix, areas: coupled } of COUPLINGS) {
      if (file.startsWith(prefix)) for (const area of coupled) areas.add(area);
    }
  }
  return areas;
};

/**
 * The property files to rerun for a change: its areas' files, any changed
 * property file, and, when the pinned seeds change, the files they name.
 */
export const selectPropertyFiles = (
  changed: readonly string[],
  all: readonly PropertyFile[],
  pinnedFiles: readonly string[] = [],
): PropertyFile[] => {
  const areas = touchedAreas(changed);
  const named = new Set(changed);
  if (named.has(PROPERTY_SEEDS_FILE)) for (const file of pinnedFiles) named.add(file);
  return all.filter(({ area, file }) => areas.has(area) || named.has(file));
};

/** The test files test/property-seeds.json names. */
const pinnedSeedFiles = (): string[] =>
  Object.keys(JSON.parse(readFileSync(path.join(REPO_ROOT, PROPERTY_SEEDS_FILE), "utf8")) as object)
    .filter((key) => !key.startsWith("$"))
    .map((key) => key.split("::")[0] as string);

const parseArgs = (argv: readonly string[]) => {
  let base = "origin/main";
  let factor: number | undefined;
  let dryRun = false;
  let all = false;
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--base") base = argv[++index] ?? base;
    else if (arg === "--factor") factor = Number(argv[++index]);
    else if (arg === "--dry-run") dryRun = true;
    else if (arg === "--all") all = true;
    else throw new Error(`Unknown argument ${String(arg)}`);
  }
  const envFactor = process.env["PROPERTY_TEST_NUM_RUNS_FACTOR"];
  const defaultFactor = all ? 1 : 5;
  factor ??= envFactor === undefined || envFactor === "" ? defaultFactor : Number(envFactor);
  if (!Number.isFinite(factor) || factor < 1) throw new Error("the factor must be a number ≥ 1");
  return { base, factor, dryRun, all };
};

const changedFiles = async (base: string): Promise<string[]> => {
  const mergeBase = (await $`git merge-base ${base} HEAD`.cwd(REPO_ROOT).text()).trim();
  const diff = await $`git diff --name-only ${mergeBase} HEAD`.cwd(REPO_ROOT).text();
  return diff.split("\n").filter((line) => line !== "");
};

if (import.meta.main) {
  const { base, factor, dryRun, all } = parseArgs(process.argv.slice(2));
  let selected: PropertyFile[];
  if (all) {
    selected = propertyFiles();
    console.log(`${String(selected.length)} property files at factor ${String(factor)}`);
  } else {
    const changed = await changedFiles(base);
    selected = selectPropertyFiles(changed, propertyFiles(), pinnedSeedFiles());
    const areas = [...new Set(selected.map(({ area }) => area))];
    console.log(
      `${String(changed.length)} changed files vs ${base}; ${String(selected.length)} property files in ${String(areas.length)} areas at factor ${String(factor)}${areas.length > 0 ? `: ${areas.join(", ")}` : ""}`,
    );
  }
  if (dryRun || selected.length === 0) {
    for (const { file } of selected) console.log(`  ${file}`);
  } else {
    const byPackage = new Map<string, PropertyFile[]>();
    for (const entry of selected) {
      byPackage.set(entry.packageDir, [...(byPackage.get(entry.packageDir) ?? []), entry]);
    }
    // Packages run side by side (at factor 10 docx-core's operation properties
    // alone take as long as all of core's); each one's output is printed whole
    // when it finishes, so a log still reads one package, one file at a time.
    const exitCodes = await Promise.all(
      [...byPackage].map(async ([packageDir, files]) => {
        const relative = files.map(({ file }) => path.relative(packageDir, file));
        const started = performance.now();
        const run = await $`bun test ${relative} 2>&1`
          .cwd(path.join(REPO_ROOT, packageDir))
          .env({ ...process.env, PROPERTY_TEST_NUM_RUNS_FACTOR: String(factor) })
          .quiet()
          .nothrow();
        const seconds = ((performance.now() - started) / 1000).toFixed(1);
        process.stdout.write(run.stdout);
        console.log(
          `${packageDir}: ${String(files.length)} files in ${seconds}s, exit ${String(run.exitCode)}`,
        );
        return run.exitCode;
      }),
    );
    // Set the code rather than exiting: output written to a pipe is flushed
    // asynchronously, and exiting here cut a failing run's log short.
    process.exitCode = exitCodes.every((code) => code === 0) ? 0 : 1;
  }
}
