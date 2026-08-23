#!/usr/bin/env bun

import { readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { Result, TaggedError } from "better-result";

import { keyDifferences, validateExactObjectKeys } from "./lib/exact-object";

const REPORT_ROOT = path.resolve(import.meta.dir, "../api-reports");
const BASELINE_PATH = path.join(import.meta.dir, "api-surface-budget.json");
const PACKAGES = ["docx-core", "core", "react", "agents", "vue", "nuxt"] as const;
const HEADROOM = 1.05;
const BYTE_FLOOR = 1024;
const LINE_FLOOR = 50;
const RATCHET_THRESHOLD = 0.95;

type PackageName = (typeof PACKAGES)[number];
type SurfaceSize = {
  reportBytes: number;
  reportLines: number;
  declarationBytes: number;
  declarationLines: number;
};
type Baseline = Record<PackageName, SurfaceSize>;

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

const measureAll = (): Baseline => ({
  "docx-core": measurePackage("docx-core"),
  core: measurePackage("core"),
  react: measurePackage("react"),
  agents: measurePackage("agents"),
  vue: measurePackage("vue"),
  nuxt: measurePackage("nuxt"),
});

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
  const docxCore = readEntry(exactBaseline.value, "docx-core");
  const core = readEntry(exactBaseline.value, "core");
  const react = readEntry(exactBaseline.value, "react");
  const agents = readEntry(exactBaseline.value, "agents");
  const vue = readEntry(exactBaseline.value, "vue");
  const nuxt = readEntry(exactBaseline.value, "nuxt");
  if (docxCore.isErr()) return docxCore;
  if (core.isErr()) return core;
  if (react.isErr()) return react;
  if (agents.isErr()) return agents;
  if (vue.isErr()) return vue;
  if (nuxt.isErr()) return nuxt;
  return Result.ok({
    "docx-core": docxCore.value,
    core: core.value,
    react: react.value,
    agents: agents.value,
    vue: vue.value,
    nuxt: nuxt.value,
  } satisfies Baseline);
};

const maximum = (baseline: number, floor: number): number =>
  Math.max(baseline * HEADROOM, baseline + floor);

const exceedsMaximum = (current: number, baseline: number, floor: number): boolean =>
  current > maximum(baseline, floor);

const ratchetAvailable = (current: number, baseline: number): boolean =>
  current < baseline * RATCHET_THRESHOLD;

const report = (measured: Baseline): void => {
  for (const name of PACKAGES) {
    const size = measured[name];
    console.log(
      `api-surface-budget: ${name} reports ${size.reportLines.toLocaleString("en-US")} lines / ` +
        `${size.reportBytes.toLocaleString("en-US")} bytes; declarations ` +
        `${size.declarationLines.toLocaleString("en-US")} lines / ` +
        `${size.declarationBytes.toLocaleString("en-US")} bytes`,
    );
  }
};

const check = (measured: Baseline): number => {
  const baselineResult = readBaseline();
  if (baselineResult.isErr()) {
    console.error(baselineResult.error.message);
    return 1;
  }
  const baseline = baselineResult.value;
  const failures: string[] = [];
  for (const name of PACKAGES) {
    for (const metric of METRIC_NAMES) {
      const floor = METRICS[metric];
      const current = measured[name][metric];
      const previous = baseline[name][metric];
      if (exceedsMaximum(current, previous, floor)) {
        failures.push(`${name}.${metric}: ${previous} -> ${current}`);
      } else if (ratchetAvailable(current, previous)) {
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

const mode = process.argv.at(2) ?? "--check";
if (process.argv.length > 3 || !["--check", "--self-test", "--write"].includes(mode)) {
  console.error("Usage: api-surface-budget.ts [--check|--self-test|--write]");
  process.exit(2);
}
if (mode === "--self-test") {
  process.exit(selfTest());
}

try {
  const measured = measureAll();
  report(measured);
  if (mode === "--write") {
    writeFileSync(BASELINE_PATH, `${JSON.stringify(measured, null, 2)}\n`);
    console.log("api-surface-budget: wrote scripts/api-surface-budget.json");
    process.exit(0);
  }
  process.exit(check(measured));
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
}
