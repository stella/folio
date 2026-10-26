#!/usr/bin/env bun
/**
 * Run the property tests of the areas a change touches at a raised run count.
 *
 * PR CI runs every property at its own `numRuns` (factor 1) inside the main
 * test step. A property that fails on a few percent of seeds can pass there and
 * redden unrelated changes later, so the manual proof reruns the property files of the
 * areas the branch touched with `PROPERTY_TEST_NUM_RUNS_FACTOR` raised (5 by
 * default): a change to `packages/core/src/compare/**` searches the compare
 * properties five times as hard, and nothing else pays for it.
 *
 * An area is `<package>/<first directory under src>` (`core/ai-edits`,
 * `core/compare`, `docx-core/markdown`, ...), plus the couplings below where
 * one directory's code is exercised by another's properties. Changed source
 * selects areas; a changed property file runs itself, and a change to
 * test/property-seeds/ runs the test file named by each changed seed file.
 *
 * Usage:
 *   bun scripts/property-areas.ts [--base <ref>] [--factor <n>] [--dry-run]
 *   bun scripts/property-areas.ts --all            (every property file)
 *   bun scripts/property-areas.ts --shard 1/4      (one partition of the selection)
 *
 * `--base` (default `origin/main`) is diffed from its merge base with HEAD.
 * `--factor` defaults to `PROPERTY_TEST_NUM_RUNS_FACTOR`, else 5 for a change
 * and 1 for `--all` (the extended proof sets the variable). `--dry-run` prints the
 * selection without running it.
 */

import { panic } from "better-result";
import { $ } from "bun";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import type { CallExpression } from "oxc-parser";

import {
  PROPERTY_SEEDS_FILE,
  testFileForSeedFile,
} from "../test/seed-registry";
import { type Node, nodeText, parseSourceModule, walkAst } from "./lib/oxc-ast";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const PACKAGES_DIR = path.join(REPO_ROOT, "packages");
export const ROOT_TEST_DIRS = ["scripts", "tests", "parity"] as const;
const PROPERTY_FILE = /\.test\.tsx?$/;

/**
 * Changed-path prefixes whose code another area's properties exercise, on top
 * of the area the path itself belongs to.
 */
export const COUPLINGS: readonly {
  prefix: string;
  areas: readonly string[];
}[] = [
  // Shared boundaries are exercised by the actual ZIP and document serializers.
  {
    prefix: "packages/utf8-boundaries/src/",
    areas: ["docx-core/zip", "core/utils"],
  },
  // The comparison's internals live under internal/compare.
  { prefix: "packages/core/src/internal/compare/", areas: ["core/compare"] },
  // Whole-story revision resolution backs accept/reject in the reviewer and
  // the editor: the tracked-resolution and headless-review properties drive it.
  {
    prefix: "packages/core/src/internal/",
    areas: ["core/ai-edits", "core/prosemirror"],
  },
  // Tracked changes are ProseMirror marks resolved by the same commands.
  { prefix: "packages/core/src/prosemirror/", areas: ["core/ai-edits"] },
  // Both markdown readers share the docx-core compiler.
  { prefix: "packages/docx-core/src/markdown/", areas: ["core/markdown"] },
  { prefix: "packages/core/src/markdown/", areas: ["docx-core/markdown"] },
  // The comparison saves and reopens through the docx reader and writer.
  { prefix: "packages/core/src/docx/serializer/", areas: ["core/compare"] },
];

type PropertyFile = {
  area: string;
  packageDir: string;
  file: string;
  weightMs: number;
};

const PROPERTY_WORKERS = 2;
type PropertyBatch = { packageDir: string; files: PropertyFile[] };

/** Split a package's serial test workload without changing its selected files. */
export const propertyBatches = (
  selected: readonly PropertyFile[],
): PropertyBatch[] => {
  const byPackage = new Map<string, PropertyFile[]>();
  for (const entry of selected) {
    const files = byPackage.get(entry.packageDir);
    if (files) files.push(entry);
    else byPackage.set(entry.packageDir, [entry]);
  }
  return [...byPackage].flatMap(([packageDir, files]) => {
    const batchCount = Math.min(PROPERTY_WORKERS, files.length);
    const batchSize = Math.ceil(files.length / batchCount);
    return Array.from({ length: batchCount }, (_, batch) => ({
      packageDir,
      files: files.slice(batch * batchSize, (batch + 1) * batchSize),
    }));
  });
};

/** Keep process concurrency bounded and retain every batch's failure status. */
export const runPropertyBatches = async (
  batches: readonly PropertyBatch[],
  run: (batch: PropertyBatch) => Promise<number>,
): Promise<number[]> => {
  const queue = batches.values();
  const exitCodes: number[] = [];
  const worker = async () => {
    for (const batch of queue) exitCodes.push(await run(batch));
  };
  await Promise.all(
    Array.from({ length: Math.min(PROPERTY_WORKERS, batches.length) }, worker),
  );
  return exitCodes;
};

const walk = (dir: string): string[] =>
  readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    if (entry.name === "node_modules" || entry.name.startsWith(".")) return [];
    const absolute = path.join(dir, entry.name);
    return entry.isDirectory() ? walk(absolute) : [absolute];
  });

// A scheduling estimate, not a measured duration or a changed test timeout.
// Synchronous properties without a declared budget use Bun's 5-second default.
const DEFAULT_PROPERTY_WEIGHT_MS = 5_000;

const callName = (expression: Node): string | undefined => {
  if (expression.type === "Identifier") return expression.name;
  if (expression.type === "MemberExpression" && !expression.computed)
    return callName(expression.object);
  if (expression.type === "CallExpression") return callName(expression.callee);
  return undefined;
};

const enclosingTest = (ancestors: readonly Node[]): Node | undefined =>
  ancestors.findLast(
    (ancestor) =>
      ancestor.type === "CallExpression" &&
      ["test", "it"].includes(callName(ancestor.callee) ?? ""),
  );

type PropertyFileProfileOptions = { file: string; text: string };

/** Derive scheduling weight from actual property drivers and their stated base budgets. */
export const propertyFileProfile = ({
  file,
  text,
}: PropertyFileProfileOptions) => {
  const { program } = parseSourceModule(file, text);
  const constants = new Map<string, Node[]>();
  const drivers: Node[] = [];
  const budgets: CallExpression[] = [];
  const defaults: CallExpression[] = [];
  const ancestry = new Map<Node, readonly Node[]>();
  walkAst(program, (node, ancestors) => {
    const parent = ancestors.at(-1);
    if (
      node.type === "VariableDeclarator" &&
      node.id.type === "Identifier" &&
      node.init !== null &&
      parent?.type === "VariableDeclaration" &&
      // `await using` bindings carry TypeScript's const flag as well.
      (parent.kind === "const" || parent.kind === "await using")
    ) {
      const existing = constants.get(node.id.name);
      if (existing) existing.push(node.init);
      else constants.set(node.id.name, [node.init]);
    }
    if (node.type === "CallExpression") {
      const { callee } = node;
      if (
        (callee.type === "Identifier" && callee.name === "assertProperty") ||
        (callee.type === "MemberExpression" &&
          !callee.computed &&
          callee.object.type === "Identifier" &&
          callee.object.name === "fc" &&
          callee.property.type === "Identifier" &&
          (callee.property.name === "assert" ||
            callee.property.name === "check"))
      ) {
        drivers.push(node);
        ancestry.set(node, [...ancestors]);
      }
      if (
        callee.type === "Identifier" &&
        callee.name === "propertyTestTimeout"
      ) {
        budgets.push(node);
        ancestry.set(node, [...ancestors]);
      }
      if (callee.type === "Identifier" && callee.name === "setDefaultTimeout")
        defaults.push(node);
    }
  });
  if (drivers.length === 0)
    return { drivers: 0, weightMs: 0, missingBudgets: 0 };

  const numericBudget = (
    expression: Node,
    resolving = new Set<string>(),
  ): number => {
    let value: number;
    const literal = expression.type === "Literal" ? expression.value : null;
    if (typeof literal === "number") value = literal;
    else if (expression.type === "ParenthesizedExpression")
      return numericBudget(expression.expression, resolving);
    else if (expression.type === "Identifier") {
      const definitions = constants.get(expression.name);
      const definition =
        definitions?.length === 1 ? definitions.at(0) : undefined;
      if (!definition || resolving.has(expression.name))
        return panic(
          `Cannot derive property budget in ${file}: ${nodeText(text, expression)}`,
        );
      return numericBudget(
        definition,
        new Set([...resolving, expression.name]),
      );
    } else {
      return panic(
        `Cannot derive property budget in ${file}: ${nodeText(text, expression)}`,
      );
    }
    if (!Number.isFinite(value) || value <= 0)
      return panic(
        `Property budget must be finite and positive in ${file}: ${String(value)}`,
      );
    return value;
  };
  let missingBudgets = 0;
  const timeoutBudget = (call: CallExpression): number => {
    const argument = call.arguments.at(0);
    if (argument === undefined) {
      missingBudgets += 1;
      return DEFAULT_PROPERTY_WEIGHT_MS;
    }
    return numericBudget(argument);
  };
  const stated = new Map<Node, number>();
  // Validate every declared helper call, including calls not used by a property.
  for (const call of budgets) stated.set(call, timeoutBudget(call));
  let defaultBudget = DEFAULT_PROPERTY_WEIGHT_MS;
  for (const call of defaults) {
    const argument = call.arguments.at(0);
    if (argument === undefined)
      return panic(`Missing default property budget in ${file}`);
    defaultBudget =
      argument.type === "CallExpression" && stated.has(argument)
        ? (stated.get(argument) ??
          panic(`Missing profiled property budget in ${file}`))
        : numericBudget(argument);
  }
  const testBudgets = new Map<Node, number>();
  for (const [call, budget] of stated) {
    const owner = enclosingTest(ancestry.get(call) ?? []);
    if (owner !== undefined) {
      if (testBudgets.has(owner))
        return panic(`Multiple property budgets in one test in ${file}`);
      testBudgets.set(owner, budget);
    }
  }
  let weightMs = 0;
  for (const driver of drivers) {
    const owner = enclosingTest(ancestry.get(driver) ?? []);
    weightMs +=
      owner === undefined
        ? defaultBudget
        : (testBudgets.get(owner) ?? defaultBudget);
  }
  if (!Number.isFinite(weightMs) || weightMs <= 0)
    return panic(
      `Property scheduling weight must be finite and positive in ${file}`,
    );
  return { drivers: drivers.length, weightMs, missingBudgets };
};

const profileProperty = (file: string) => {
  const profile = propertyFileProfile({
    file,
    text: readFileSync(file, "utf8"),
  });
  if (profile.missingBudgets > 0)
    console.warn(
      `${file}: ${String(profile.missingBudgets)} missing property budget arguments; scheduling uses Bun's 5000ms default estimate.`,
    );
  return profile;
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
      .filter((file) => PROPERTY_FILE.test(file))
      .flatMap((file) => {
        const profile = profileProperty(file);
        if (profile.drivers === 0) return [];
        const repoPath = path.relative(REPO_ROOT, file).replaceAll("\\", "/");
        return {
          area: areaOf(repoPath) as string,
          packageDir: `packages/${pkg}`,
          file: repoPath,
          weightMs: profile.weightMs,
        };
      })
      .toSorted((a, b) => a.file.localeCompare(b.file));
  });
  const rootTests = ROOT_TEST_DIRS.flatMap((root) =>
    walk(path.join(REPO_ROOT, root))
      .filter((file) => PROPERTY_FILE.test(file))
      .flatMap((file) => {
        const profile = profileProperty(file);
        if (profile.drivers === 0) return [];
        return {
          area: root,
          packageDir: ".",
          file: path.relative(REPO_ROOT, file).replaceAll("\\", "/"),
          weightMs: profile.weightMs,
        };
      }),
  );
  return [...packages, ...rootTests].toSorted((a, b) =>
    a.file.localeCompare(b.file),
  );
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
): PropertyFile[] => {
  const areas = touchedAreas(changed);
  const named = new Set(changed);
  for (const file of changed) {
    if (file.startsWith(`${PROPERTY_SEEDS_FILE}/`))
      named.add(testFileForSeedFile(file));
  }
  return all.filter(({ area, file }) => areas.has(area) || named.has(file));
};

type PropertyShard = { index: number; total: number };

/** Longest-first greedy bin packing retains the complete selection and every property run. */
export const shardPropertyFiles = (
  files: readonly PropertyFile[],
  { index, total }: PropertyShard,
): PropertyFile[] => {
  if (
    !Number.isSafeInteger(total) ||
    total < 1 ||
    !Number.isSafeInteger(index) ||
    index < 1 ||
    index > total
  )
    return panic(
      "Property shard must be an integer index/total with 1 ≤ index ≤ total.",
    );
  const shards = Array.from({ length: total }, () => ({
    files: files.slice(0, 0),
    weightMs: 0,
  }));
  for (const file of files.toSorted(
    (a, b) => b.weightMs - a.weightMs || a.file.localeCompare(b.file),
  )) {
    if (!Number.isFinite(file.weightMs) || file.weightMs <= 0)
      return panic(
        `Property scheduling weight must be finite and positive: ${file.file}`,
      );
    let target = shards.at(0) ?? panic("Property scheduling requires a shard");
    for (const shard of shards)
      if (shard.weightMs < target.weightMs) target = shard;
    target.files.push(file);
    target.weightMs += file.weightMs;
  }
  return (
    shards.at(index - 1) ?? panic("Requested property shard is unavailable")
  ).files.toSorted((a, b) => a.file.localeCompare(b.file));
};

const parseArgs = (argv: readonly string[]) => {
  let base = "origin/main";
  let factor: number | undefined;
  let dryRun = false;
  let all = false;
  let shard: PropertyShard | undefined;
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--base") base = argv[++index] ?? base;
    else if (arg === "--factor") factor = Number(argv[++index]);
    else if (arg === "--dry-run") dryRun = true;
    else if (arg === "--all") all = true;
    else if (arg === "--shard") {
      const parts = (argv[++index] ?? "").split("/");
      if (parts.length !== 2)
        return panic("Property shard must use index/total.");
      shard = { index: Number(parts.at(0)), total: Number(parts.at(1)) };
    } else throw new Error(`Unknown argument ${String(arg)}`);
  }
  const envFactor = process.env["PROPERTY_TEST_NUM_RUNS_FACTOR"];
  const defaultFactor = all ? 1 : 5;
  factor ??=
    envFactor === undefined || envFactor === ""
      ? defaultFactor
      : Number(envFactor);
  if (!Number.isFinite(factor) || factor < 1)
    throw new Error("the factor must be a number ≥ 1");
  return { base, factor, dryRun, all, shard };
};

const changedFiles = async (base: string): Promise<string[]> => {
  const mergeBase = (
    await $`git merge-base ${base} HEAD`.cwd(REPO_ROOT).text()
  ).trim();
  const diff = await $`git diff --name-only ${mergeBase} HEAD`
    .cwd(REPO_ROOT)
    .text();
  return diff.split("\n").filter((line) => line !== "");
};

if (import.meta.main) {
  const { base, factor, dryRun, all, shard } = parseArgs(process.argv.slice(2));
  let selected: PropertyFile[];
  if (all) {
    selected = propertyFiles();
    console.log(
      `${String(selected.length)} property files at factor ${String(factor)}`,
    );
  } else {
    const changed = await changedFiles(base);
    selected = selectPropertyFiles(changed, propertyFiles());
    const areas = [...new Set(selected.map(({ area }) => area))];
    console.log(
      `${String(changed.length)} changed files vs ${base}; ${String(selected.length)} property files in ${String(areas.length)} areas at factor ${String(factor)}${areas.length > 0 ? `: ${areas.join(", ")}` : ""}`,
    );
  }
  if (shard !== undefined) {
    selected = shardPropertyFiles(selected, shard);
    console.log(
      `Shard ${String(shard.index)}/${String(shard.total)}: ${String(selected.length)} property files`,
    );
  }
  if (dryRun || selected.length === 0) {
    for (const { file } of selected) console.log(`  ${file}`);
  } else {
    // Two bounded workers can share one large package's property workload.
    // Stream progress and counterexamples so cancellation retains diagnostics.
    const exitCodes = await runPropertyBatches(
      propertyBatches(selected),
      async ({ packageDir, files }) => {
        const relative = files.map(({ file }) =>
          path.relative(packageDir, file),
        );
        const started = performance.now();
        console.log(
          `${packageDir}: starting ${String(files.length)} files: ${relative.join(", ")}`,
        );
        const run = await $`bun test ${relative} 2>&1`
          .cwd(path.join(REPO_ROOT, packageDir))
          .env({
            ...process.env,
            PROPERTY_TEST_NUM_RUNS_FACTOR: String(factor),
          })
          .nothrow();
        const seconds = ((performance.now() - started) / 1000).toFixed(1);
        console.log(
          `${packageDir}: ${String(files.length)} files in ${seconds}s, exit ${String(run.exitCode)}`,
        );
        return run.exitCode;
      },
    );
    // Set the code rather than exiting: output written to a pipe is flushed
    // asynchronously, and exiting here cut a failing run's log short.
    process.exitCode = exitCodes.every((code) => code === 0) ? 0 : 1;
  }
}
