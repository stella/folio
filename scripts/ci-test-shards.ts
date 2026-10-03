#!/usr/bin/env bun
// Refresh ci-test-timings.json from per-file group spans in a successful full-suite log.
// Keep millisecond durations above one second; unmeasured files still run.
import { panic } from "better-result";
import { existsSync, readFileSync } from "node:fs";
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
export const FAST_SIBLING_TEST_CAP = 150;
const SOURCE_FILE = /^packages\/[^/]+\/src\/.*\.[cm]?[jt]sx?$/u;

/** Deleted paths never expand a fast suite; git returns rename destinations. */
export const changedTestPaths = (base: string, repoRoot = REPO_ROOT): string[] => {
  const result = Bun.spawnSync(
    ["git", "diff", "--name-only", "--diff-filter=ACMRT", "-z", `${base}...HEAD`, "--"],
    { cwd: repoRoot, stdout: "pipe", stderr: "pipe" },
  );
  if (result.exitCode !== 0) panic(`Cannot select changed tests: ${result.stderr.toString()}`);
  return result.stdout
    .toString()
    .split("\0")
    .filter((file) => file !== "" && existsSync(path.join(repoRoot, file)));
};

type FastTestSuitesOptions = {
  suites: TestSuite[];
  changedFiles: string[];
  focusedFiles?: string[];
};

/** Extend the fixed list within the owning test commands, retaining their preloads. */
export const selectFastTestSuites = ({
  suites,
  changedFiles,
  focusedFiles = focusedTests,
}: FastTestSuitesOptions): TestSuite[] => {
  const changed = new Set(changedFiles);
  const siblingDirectories = new Set(
    changedFiles
      .filter((file) => SOURCE_FILE.test(file) && !TEST_FILE.test(file))
      .map((file) => path.posix.dirname(file)),
  );
  const discovered = new Set(suites.flatMap(({ files }) => files));
  const focused = new Set(
    focusedFiles.filter((file) => discovered.has(file) || existsSync(path.join(REPO_ROOT, file))),
  );
  const selected = suites.map((suite) => {
    const siblings = new Set(
      suite.files.filter((file) => siblingDirectories.has(path.posix.dirname(file))),
    );
    const files =
      suite.cwd.startsWith("packages/") && siblings.size > FAST_SIBLING_TEST_CAP
        ? suite.files
        : suite.files.filter((file) => changed.has(file) || siblings.has(file));
    return { ...suite, files: files.filter((file) => !focused.has(file)) };
  });
  // Preserve the existing fixed-list command and avoid running its files twice.
  selected.push({ cwd: ".", preloads: [], files: [...focused] });
  return selected.filter(({ files }) => files.length > 0);
};

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
  const base = process.env["CI_TEST_BASE"];
  const selectedSuites =
    depth === "full"
      ? suites
      : selectFastTestSuites({
          suites,
          changedFiles: base ? changedTestPaths(base) : [],
        });
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
