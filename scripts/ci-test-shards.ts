#!/usr/bin/env bun
// Refresh ci-test-timings.json from per-file group spans in a successful full-suite log.
// Keep millisecond durations above one second; unmeasured files still run.
import { panic } from "better-result";
import { readFileSync } from "node:fs";
import path from "node:path";
import focusedTests from "./ci-focused-tests.json";
import timings from "./ci-test-timings.json";

export const TEST_SHARD_COUNT = 4;
const REPO_ROOT = path.resolve(import.meta.dir, "..");
const TEST_FILE = /(?:\.test|\.spec|_test|_spec)\.(?:[cm]?[jt]sx?)$/u;

type TestSuite = { cwd: string; preloads: string[]; files: string[] };

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
  for (const manifest of new Bun.Glob("packages/*/package.json").scanSync({ cwd: repoRoot })) {
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
    suites.push({ cwd, preloads, files: discoverFiles(cwd, roots) });
  }
  for (const root of ["scripts", "benchmarks/compare"]) {
    suites.push({ cwd: ".", preloads: [], files: discoverFiles(".", [root]) });
  }
  return suites.toSorted((a, b) => a.cwd.localeCompare(b.cwd));
};

// Missing measurements get a small scheduling estimate, never exclusion.
export const assignTestShards = (files: string[]) => {
  const measured = new Map(Object.entries(timings));
  const weight = (file: string) => measured.get(file) ?? 100;
  const loads = Array.from({ length: TEST_SHARD_COUNT }, () => 0);
  const shards = Array.from({ length: TEST_SHARD_COUNT }, () => new Set<string>());
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

export const focusedTestFiles = focusedTests;

if (import.meta.main) {
  const args = process.argv.slice(2);
  const shard = Number(args.at(args.indexOf("--shard") + 1));
  const depth = args.at(args.indexOf("--depth") + 1);
  if (
    !Number.isInteger(shard) ||
    shard < 1 ||
    shard > TEST_SHARD_COUNT ||
    (depth !== "fast" && depth !== "full")
  ) {
    panic("Usage: bun scripts/ci-test-shards.ts --shard <1..4> --depth <fast|full>");
  }
  const suites = discoverTestSuites();
  const allFiles = [...suites.flatMap(({ files }) => files), ...focusedTests];
  const assigned = assignTestShards(allFiles).at(shard - 1);
  if (!assigned) panic("Test shard unavailable");
  const selectedSuites =
    depth === "full" ? suites : [{ cwd: ".", preloads: [], files: focusedTests }];
  let exitCode = 0;
  for (const { cwd, preloads, files } of selectedSuites) {
    const selected = files.filter((file) => assigned.has(file));
    if (selected.length === 0) continue;
    console.log(`Shard ${shard}/${TEST_SHARD_COUNT}: ${cwd} (${selected.length} files)`);
    const result = Bun.spawnSync(
      [
        "bun",
        "test",
        ...preloads.flatMap((preload) => ["--preload", preload]),
        ...selected.map((file) => `./${path.posix.relative(cwd, file)}`),
      ],
      { cwd: path.join(REPO_ROOT, cwd), stdout: "inherit", stderr: "inherit" },
    );
    if (result.exitCode !== 0) exitCode = 1;
  }
  process.exit(exitCode);
}
