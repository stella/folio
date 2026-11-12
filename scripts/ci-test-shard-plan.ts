#!/usr/bin/env bun
// Plans the CI test shards from measured per-file timings: the shard count is
// the smallest whose longest shard fits the budget, so the workflow's shard
// matrix follows the measurements instead of a hand-kept job list.
import { panic } from "better-result";
import { appendFileSync, existsSync, readFileSync } from "node:fs";
import path from "node:path";

import focusedTests from "./ci-focused-tests.json";
import timings from "./ci-test-timings.json";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const TEST_FILE = /(?:\.test|\.spec|_test|_spec)\.(?:[cm]?[jt]sx?)$/u;

/**
 * Planned measured time per shard. Measured file times leave out process
 * start, preloads and runner variance, so a shard's job runs longer than its
 * plan; the budget keeps that inside the job's timeout-minutes.
 */
export const SHARD_BUDGET_MS = 9 * 60_000;

export type TestSuite = { cwd: string; preloads: string[]; files: string[] };
export type TestTimings = Record<string, number>;

const median = (values: number[]) => {
  const ordered = values.toSorted((a, b) => a - b);
  const middle = Math.floor(ordered.length / 2);
  if (ordered.length === 0) return 100;
  if (ordered.length % 2 === 1) return ordered.at(middle) ?? 100;
  return Math.round(
    ((ordered.at(middle - 1) ?? 0) + (ordered.at(middle) ?? 0)) / 2,
  );
};

export const discoverTestSuites = (repoRoot = REPO_ROOT): TestSuite[] => {
  const discoverFiles = (cwd: string, roots: string[]) => {
    const files = new Set<string>();
    for (const root of roots) {
      for (const file of new Bun.Glob(`${root}/**/*`).scanSync({
        cwd: path.join(repoRoot, cwd),
        onlyFiles: true,
      })) {
        if (TEST_FILE.test(file)) files.add(path.posix.join(cwd, file));
      }
    }
    return [...files].toSorted();
  };

  const suites: TestSuite[] = [];
  // The root command runs every packages/* workspace with a test script.
  for (const manifest of new Bun.Glob("packages/*/package.json").scanSync({
    cwd: repoRoot,
  })) {
    const pkg = JSON.parse(readFileSync(path.join(repoRoot, manifest), "utf8"));
    const command = pkg.scripts?.test;
    if (command === undefined) continue;
    if (typeof command !== "string" || !command.startsWith("bun test ")) {
      panic(`Unsupported test command in ${manifest}: ${command}`);
    }
    const cwd = path.dirname(manifest);
    const args = command.slice("bun test ".length).split(/\s+/u);
    const preloads: string[] = [];
    const roots: string[] = [];
    for (let i = 0; i < args.length; i++) {
      const arg = args.at(i);
      if (arg === "--preload") {
        const preload = args.at(++i);
        if (!preload) panic(`Missing preload in ${manifest}`);
        preloads.push(preload);
        continue;
      }
      if (!arg || arg.startsWith("-") || !/^[\w./-]+$/u.test(arg)) {
        panic(`Unsupported test argument in ${manifest}: ${arg}`);
      }
      roots.push(arg);
    }
    const bunfig = path.join(repoRoot, cwd, "bunfig.toml");
    if (existsSync(bunfig)) {
      const config: unknown = Bun.TOML.parse(readFileSync(bunfig, "utf8"));
      if (typeof config === "object" && config && "test" in config) {
        const test = config.test;
        if (typeof test === "object" && test && "preload" in test) {
          const configured = test.preload;
          if (
            !Array.isArray(configured) ||
            configured.some((preload) => typeof preload !== "string")
          ) {
            panic(
              `Unsupported test preload in ${path.posix.join(cwd, "bunfig.toml")}`,
            );
          }
          preloads.push(...configured);
        }
      }
    }
    suites.push({ cwd, preloads, files: discoverFiles(cwd, roots) });
  }
  for (const root of ["scripts", "benchmarks/compare"]) {
    suites.push({ cwd: ".", preloads: [], files: discoverFiles(".", [root]) });
  }
  return suites.toSorted((a, b) => a.cwd.localeCompare(b.cwd));
};

// Missing measurements get a small scheduling estimate, never exclusion.
export const assignTestShards = (
  files: string[],
  count: number,
  measurements: TestTimings = timings,
) => {
  if (!Number.isInteger(count) || count < 1)
    panic(`Invalid test shard count ${count}`);
  const weight = weigh(measurements);
  const loads = Array.from({ length: count }, () => 0);
  const shards = Array.from({ length: count }, () => new Set<string>());
  const ordered = [...new Set(files)].toSorted(
    (a, b) => weight(b) - weight(a) || a.localeCompare(b),
  );
  for (const file of ordered) {
    const shard = loads.indexOf(Math.min(...loads));
    const target = shards.at(shard);
    const load = loads.at(shard);
    if (!target || load === undefined) panic("Test shard unavailable");
    target.add(file);
    loads[shard] = load + weight(file);
  }
  return shards;
};

export const shardLoads = (
  shards: Set<string>[],
  measurements: TestTimings = timings,
) => {
  const weight = weigh(measurements);
  return shards.map((shard) =>
    [...shard].reduce((total, file) => total + weight(file), 0),
  );
};

const weigh = (measurements: TestTimings) => {
  const fallback = median(Object.values(measurements));
  return (file: string) => measurements[file] ?? fallback;
};

/** Every file a full run executes, in discovery order. */
export const plannedTestFiles = (repoRoot = REPO_ROOT): string[] => [
  ...discoverTestSuites(repoRoot).flatMap(({ files }) => files),
  ...focusedTests,
];

/** The fewest shards whose longest planned load fits the budget. */
export const planTestShards = (
  files: string[],
  measurements: TestTimings = timings,
  budget = SHARD_BUDGET_MS,
) => {
  const unique = [...new Set(files)];
  const weight = weigh(measurements);
  const oversized = unique.filter((file) => weight(file) > budget);
  if (oversized.length > 0)
    panic(
      `Test files exceed the ${budget} ms shard budget on their own: ${oversized.join(", ")}`,
    );
  const total = unique.reduce((sum, file) => sum + weight(file), 0);
  for (
    let count = Math.max(1, Math.ceil(total / budget));
    count <= Math.max(1, unique.length);
    count++
  ) {
    const shards = assignTestShards(unique, count, measurements);
    if (Math.max(...shardLoads(shards, measurements)) <= budget) return shards;
  }
  return panic("No shard count fits the test budget");
};

if (import.meta.main) {
  const shards = planTestShards(plannedTestFiles());
  const loads = shardLoads(shards);
  console.log(
    `Planned ${shards.length} test shards: ${loads.map((load, index) => `${index + 1}=${load}ms`).join(", ")}`,
  );
  const output = process.env["GITHUB_OUTPUT"];
  const matrix = JSON.stringify(shards.map((_, index) => index + 1));
  if (output) appendFileSync(output, `test_shards=${matrix}\n`);
  else console.log(`test_shards=${matrix}`);
}
