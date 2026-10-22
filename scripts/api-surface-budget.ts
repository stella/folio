#!/usr/bin/env bun

import { mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import { Result, TaggedError, panic } from "better-result";

import { keyDifferences, validateExactObjectKeys } from "./lib/exact-object";
import { PUBLISHED_PACKAGES } from "./lib/published-packages";
import { installFrozenRepositoryDependencies } from "./clean-room-install";

const REPORT_ROOT = path.resolve(import.meta.dir, "../api-reports");
const BASELINE_PATH = path.join(import.meta.dir, "api-surface-budget.json");
const PACKAGES = PUBLISHED_PACKAGES.map(({ slug }) => slug);
const HEADROOM = 1.05;
const BYTE_FLOOR = 1024;
const LINE_FLOOR = 50;
const RATCHET_THRESHOLD = 0.95;

type PackageName = (typeof PACKAGES)[number];
export type SurfaceSize = {
  reportBytes: number;
  reportLines: number;
  declarationBytes: number;
  declarationLines: number;
};
type Baseline = Map<PackageName, SurfaceSize>;
export type SurfaceBudget = ReadonlyMap<string, SurfaceSize>;

const METRIC_NAMES = [
  "reportBytes",
  "reportLines",
  "declarationBytes",
  "declarationLines",
] as const satisfies readonly (keyof SurfaceSize)[];
const METRICS = {
  reportBytes: BYTE_FLOOR,
  reportLines: LINE_FLOOR,
  declarationBytes: BYTE_FLOOR,
  declarationLines: LINE_FLOOR,
} as const satisfies Record<keyof SurfaceSize, number>;

const isSize = (value: unknown): value is number =>
  typeof value === "number" && Number.isFinite(value) && value >= 0;

class ApiSurfaceBudgetError extends TaggedError("ApiSurfaceBudgetError")<{
  message: string;
  cause?: unknown;
}> {}

const rejectsExactKeys = (value: unknown, expected: readonly string[]): boolean =>
  validateExactObjectKeys(value, expected, "self-test baseline").isErr();

type FileSize = { bytes: number; lines: number };

const measureFiles = (directory: string, files: readonly string[]): FileSize => {
  let bytes = 0;
  let lines = 0;
  for (const file of files) {
    const contents = readFileSync(path.join(directory, file), "utf8");
    bytes += Buffer.byteLength(contents);
    lines += contents.split("\n").length - 1;
  }
  return { bytes, lines };
};

const measurePackage = (name: PackageName): SurfaceSize => {
  const reportDirectory = path.join(REPORT_ROOT, name);
  const reportFiles = readdirSync(reportDirectory)
    .filter((file) => file.endsWith(".api.md"))
    .toSorted();
  if (reportFiles.length === 0) {
    throw new Error(`No API reports found for ${name}`);
  }
  const distDirectory = path.resolve(REPORT_ROOT, "../packages", name, "dist");
  const declarationFiles = readdirSync(distDirectory, { recursive: true })
    .filter(
      (file): file is string =>
        typeof file === "string" &&
        (file.endsWith(".d.ts") || file.endsWith(".d.mts") || file.endsWith(".d.cts")),
    )
    .toSorted();
  if (declarationFiles.length === 0) {
    throw new Error(`No built declarations found for ${name}`);
  }
  const reports = measureFiles(reportDirectory, reportFiles);
  const declarations = measureFiles(distDirectory, declarationFiles);
  return {
    reportBytes: reports.bytes,
    reportLines: reports.lines,
    declarationBytes: declarations.bytes,
    declarationLines: declarations.lines,
  };
};

const measureAll = (): Baseline => new Map(PACKAGES.map((name) => [name, measurePackage(name)]));

const sizeFor = (baseline: Baseline, name: PackageName): SurfaceSize =>
  baseline.get(name) ?? panic("Missing measured API surface", { name });

const readEntry = (value: Record<string, unknown>, name: PackageName) => {
  if (typeof value !== "object" || value === null || !(name in value)) {
    return Result.err(new ApiSurfaceBudgetError({ message: `Missing ${name} API surface budget` }));
  }
  const entry = value[name];
  if (typeof entry !== "object" || entry === null) {
    return Result.err(new ApiSurfaceBudgetError({ message: `Invalid ${name} API surface budget` }));
  }
  const exactEntry = validateExactObjectKeys(entry, METRIC_NAMES, `${name} API surface budget`);
  if (exactEntry.isErr()) {
    return exactEntry;
  }
  const reportBytes = exactEntry.value["reportBytes"];
  const reportLines = exactEntry.value["reportLines"];
  const declarationBytes = exactEntry.value["declarationBytes"];
  const declarationLines = exactEntry.value["declarationLines"];
  if (
    !isSize(reportBytes) ||
    !isSize(reportLines) ||
    !isSize(declarationBytes) ||
    !isSize(declarationLines)
  ) {
    return Result.err(new ApiSurfaceBudgetError({ message: `Invalid ${name} API surface budget` }));
  }
  return Result.ok({ reportBytes, reportLines, declarationBytes, declarationLines });
};

const readBaseline = () => {
  const parseJson = (): unknown => JSON.parse(readFileSync(BASELINE_PATH, "utf8"));
  const parsed = Result.try({
    try: parseJson,
    catch: (cause) =>
      new ApiSurfaceBudgetError({
        message: "Invalid scripts/api-surface-budget.json: expected JSON",
        cause,
      }),
  });
  if (parsed.isErr()) {
    return parsed;
  }
  const exactBaseline = validateExactObjectKeys(
    parsed.value,
    PACKAGES,
    "scripts/api-surface-budget.json",
  );
  if (exactBaseline.isErr()) {
    return exactBaseline;
  }
  const baseline: Baseline = new Map();
  for (const name of PACKAGES) {
    const entry = readEntry(exactBaseline.value, name);
    if (entry.isErr()) return entry;
    baseline.set(name, entry.value);
  }
  return Result.ok(baseline);
};

const maximum = (baseline: number, floor: number): number =>
  Math.max(baseline * HEADROOM, baseline + floor);

const exceedsMaximum = (current: number, baseline: number, floor: number): boolean =>
  current > maximum(baseline, floor);

const ratchetAvailable = (current: number, baseline: number): boolean =>
  current < baseline * RATCHET_THRESHOLD;

const report = (measured: Baseline): void => {
  for (const name of PACKAGES) {
    const size = sizeFor(measured, name);
    console.log(
      `api-surface-budget: ${name} reports ${size.reportLines.toLocaleString("en-US")} lines / ` +
        `${size.reportBytes.toLocaleString("en-US")} bytes; declarations ` +
        `${size.declarationLines.toLocaleString("en-US")} lines / ` +
        `${size.declarationBytes.toLocaleString("en-US")} bytes`,
    );
  }
};

type BudgetEvaluation = { failures: string[]; warnings: string[] };

export const evaluateBudget = (
  base: SurfaceBudget,
  measured: SurfaceBudget,
  budget: SurfaceBudget,
): BudgetEvaluation => {
  const failures: string[] = [];
  const warnings: string[] = [];
  for (const [name, size] of measured) {
    const baseSize = base.get(name) ?? panic("Missing base API surface", { name });
    const maximumSize = budget.get(name) ?? panic("Missing API surface budget", { name });
    for (const metric of METRIC_NAMES) {
      const current = size[metric];
      const previous = baseSize[metric];
      const budgetMaximum = maximumSize[metric];
      if (current <= budgetMaximum) continue;
      const detail = `${name}.${metric}: base ${previous}, new ${current}, budget ${budgetMaximum}`;
      if (current > previous) {
        failures.push(
          `${detail}; bump scripts/api-surface-budget.json ${name}.${metric} to at least ${current}`,
        );
      } else {
        warnings.push(`${detail}; existing base drift did not grow`);
      }
    }
  }
  return { failures, warnings };
};

const check = (base: Baseline, measured: Baseline): number => {
  const baselineResult = readBaseline();
  if (baselineResult.isErr()) {
    console.error(baselineResult.error.message);
    return 1;
  }
  const baseline = baselineResult.value;
  const { failures, warnings } = evaluateBudget(base, measured, baseline);
  for (const warning of warnings) console.warn(`api-surface-budget: warning: ${warning}`);
  for (const name of PACKAGES) {
    for (const metric of METRIC_NAMES) {
      const current = sizeFor(measured, name)[metric];
      const previous = sizeFor(baseline, name)[metric];
      if (ratchetAvailable(current, previous)) {
        console.log(
          `api-surface-budget: ratchet available for ${name}.${metric}: ` +
            `${previous} -> ${current}; run bun run api:budget:write`,
        );
      }
    }
  }
  if (failures.length === 0) {
    console.log("api-surface-budget: all public declaration budgets are within limits");
    return 0;
  }
  console.error("api-surface-budget: public declaration surface exceeded its maximum:");
  for (const failure of failures) {
    console.error(`  ${failure}`);
  }
  return 1;
};

const measureBase = (revision: string): Baseline => {
  if (!/^[\da-f]{40}$/iu.test(revision)) throw new TypeError("API budget base must be a full SHA.");
  const root = path.resolve(import.meta.dir, "..");
  const cacheRoot = path.join(root, ".cache");
  mkdirSync(cacheRoot, { recursive: true });
  const worktree = mkdtempSync(path.join(cacheRoot, "api-budget-base-"));
  const add = Bun.spawnSync(["git", "worktree", "add", "--detach", worktree, revision], {
    cwd: root,
    stdout: "pipe",
    stderr: "pipe",
  });
  if (add.exitCode !== 0) {
    rmSync(worktree, { recursive: true, force: true });
    throw new Error(`Cannot create API budget base worktree: ${add.stderr.toString().trim()}`);
  }
  try {
    installFrozenRepositoryDependencies(worktree);
    const build = Bun.spawnSync(["bun", "run", "build"], {
      cwd: worktree,
      stdout: "inherit",
      stderr: "inherit",
    });
    if (build.exitCode !== 0) throw new Error("Cannot build the base API declaration surface.");
    const sizes = new Map<PackageName, SurfaceSize>();
    for (const name of PACKAGES) {
      const reportDirectory = path.join(worktree, "api-reports", name);
      const reports = measureFiles(
        reportDirectory,
        readdirSync(reportDirectory)
          .filter((file) => file.endsWith(".api.md"))
          .toSorted(),
      );
      const distDirectory = path.join(worktree, "packages", name, "dist");
      const declarations = measureFiles(
        distDirectory,
        readdirSync(distDirectory, { recursive: true })
          .filter(
            (file): file is string =>
              typeof file === "string" && /\.d\.(?:ts|mts|cts)$/u.test(file),
          )
          .toSorted(),
      );
      sizes.set(name, {
        reportBytes: reports.bytes,
        reportLines: reports.lines,
        declarationBytes: declarations.bytes,
        declarationLines: declarations.lines,
      });
    }
    return sizes;
  } finally {
    Bun.spawnSync(["git", "worktree", "remove", "--force", worktree], { cwd: root });
    rmSync(worktree, { recursive: true, force: true });
  }
};

const selfTest = (): number => {
  if (
    maximum(1_000_000, BYTE_FLOOR) !== 1_050_000 ||
    maximum(1_000, BYTE_FLOOR) !== 2024 ||
    exceedsMaximum(2024, 1_000, BYTE_FLOOR) ||
    !exceedsMaximum(2025, 1_000, BYTE_FLOOR) ||
    ratchetAvailable(950, 1_000) ||
    !ratchetAvailable(949, 1_000) ||
    !isSize(0) ||
    isSize(Number.NaN) ||
    isSize(-1) ||
    Object.keys(METRICS).length !== 4
  ) {
    console.error("api-surface-budget: self-test failed");
    return 1;
  }
  if (
    rejectsExactKeys({ core: {} }, ["core"]) ||
    !rejectsExactKeys({}, ["core"]) ||
    !rejectsExactKeys({ core: {}, stale: {} }, ["core"])
  ) {
    console.error("api-surface-budget: baseline-key self-test failed");
    return 1;
  }

  const publishedPackageSlugs: string[] = [];
  const packagesRoot = path.resolve(REPORT_ROOT, "../packages");
  for (const directory of readdirSync(packagesRoot, { withFileTypes: true })) {
    if (!directory.isDirectory()) {
      continue;
    }
    const packageJson: unknown = JSON.parse(
      readFileSync(path.join(packagesRoot, directory.name, "package.json"), "utf8"),
    );
    if (
      typeof packageJson === "object" &&
      packageJson !== null &&
      !("private" in packageJson && packageJson.private === true)
    ) {
      publishedPackageSlugs.push(directory.name);
    }
  }
  const packageCoverage = keyDifferences(publishedPackageSlugs, PACKAGES);
  if (packageCoverage.missing.length > 0 || packageCoverage.extra.length > 0) {
    console.error(
      "api-surface-budget: published-package coverage self-test failed; " +
        `missing: ${packageCoverage.missing.join(", ") || "none"}; ` +
        `extra: ${packageCoverage.extra.join(", ") || "none"}`,
    );
    return 1;
  }
  console.log("api-surface-budget: self-test passed");
  return 0;
};

if (import.meta.main) {
  const args = process.argv.slice(2);
  const mode = args.at(0) ?? "--check";
  const baseIndex = args.indexOf("--base");
  const baseArgument = baseIndex === -1 ? undefined : args.at(baseIndex + 1);
  const base = baseArgument ?? process.env["API_BASE"];
  const expectedLength = baseArgument === undefined ? 1 : 3;
  if (
    args.length > expectedLength ||
    !["--check", "--self-test", "--write"].includes(mode) ||
    (baseIndex !== -1 && (baseIndex !== 1 || baseArgument === undefined))
  ) {
    console.error("Usage: api-surface-budget.ts [--check|--self-test|--write] [--base <sha>]");
    process.exit(2);
  }
  if (mode === "--self-test") process.exit(selfTest());

  try {
    const measured = measureAll();
    report(measured);
    if (mode === "--write") {
      writeFileSync(BASELINE_PATH, `${JSON.stringify(Object.fromEntries(measured), null, 2)}\n`);
      console.log("api-surface-budget: wrote scripts/api-surface-budget.json");
      process.exit(0);
    }
    if (!base) throw new TypeError("API budget checks require --base <sha>.");
    process.exit(check(measureBase(base), measured));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(1);
  }
}
