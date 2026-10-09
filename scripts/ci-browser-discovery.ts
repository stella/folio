#!/usr/bin/env bun
/** Discover every tracked browser suite without browsers, builds or web servers. */
import { createRequire } from "node:module";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const CONFIG_GLOB = new Bun.Glob("**/playwright*.config.*");
const CONFIG_EXTENSION = /\.(?:ts|js|mts|cts|mjs|cjs)$/u;

type CommandOptions = { command: string[]; cwd: string };
const runCommand = ({ command, cwd }: CommandOptions) => {
  const result = Bun.spawnSync(command, { cwd, stdout: "pipe", stderr: "pipe" });
  return {
    exitCode: result.exitCode,
    stdout: new TextDecoder("utf-8", { fatal: true }).decode(result.stdout),
    stderr: new TextDecoder("utf-8", { fatal: true }).decode(result.stderr),
  };
};
type DiscoveryOptions = { repoRoot?: string; run?: typeof runCommand };
export const trackedBrowserConfigs = ({
  repoRoot = REPO_ROOT,
  run = runCommand,
}: DiscoveryOptions = {}): string[] => {
  const result = run({ command: ["git", "ls-files", "-z"], cwd: repoRoot });
  if (result.exitCode !== 0)
    throw new TypeError(`Cannot enumerate tracked Playwright configs:\n${result.stderr}`);
  if (!result.stdout.endsWith("\0"))
    throw new TypeError("Tracked-file discovery must produce a nonempty NUL-terminated list.");
  const files = result.stdout.slice(0, -1).split("\0");
  for (const file of files)
    if (
      file === "" ||
      file === ".." ||
      file === "." ||
      path.posix.isAbsolute(file) ||
      path.posix.normalize(file) !== file ||
      file.startsWith("../")
    )
      throw new TypeError(`Invalid tracked repository path: ${JSON.stringify(file)}`);
  const configs = [...new Set(files)]
    .filter((file) => CONFIG_GLOB.match(file) && CONFIG_EXTENSION.test(file))
    .toSorted();
  if (configs.length === 0) throw new TypeError("No tracked Playwright configs were discovered.");
  return configs;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const suiteTestCount = (suites: unknown[]): number => {
  let count = 0;
  for (const suite of suites) {
    if (!isRecord(suite) || !Array.isArray(suite["specs"]))
      throw new TypeError("Playwright JSON suites must contain specs arrays.");
    for (const spec of suite["specs"]) {
      if (!isRecord(spec) || !Array.isArray(spec["tests"]))
        throw new TypeError("Playwright JSON specs must contain tests arrays.");
      count += spec["tests"].length;
    }
    if (suite["suites"] === undefined) continue;
    if (!Array.isArray(suite["suites"]))
      throw new TypeError("Playwright JSON nested suites must be arrays.");
    count += suiteTestCount(suite["suites"]);
  }
  return count;
};

type BrowserDiscoveryOptions = DiscoveryOptions & {
  cliPath?: string;
  log?: (message: string) => void;
};
type BrowserConfigDiscoveryOptions = BrowserDiscoveryOptions & { config: string };

/** Discover one complete config graph; callers enumerate configs from Git. */
export const checkBrowserConfigDiscovery = ({
  config,
  repoRoot = REPO_ROOT,
  run = runCommand,
  cliPath,
  log = console.log,
}: BrowserConfigDiscoveryOptions) => {
  const playwrightCli =
    cliPath ?? createRequire(path.join(repoRoot, "package.json")).resolve("@playwright/test/cli");
  const started = performance.now();
  log(`Browser discovery: starting ${config}.`);
  const result = run({
    command: [
      "node",
      playwrightCli,
      "test",
      "--config",
      path.join(repoRoot, config),
      "--list",
      "--reporter=json",
    ],
    cwd: repoRoot,
  });
  const elapsedMs = Math.round(performance.now() - started);
  log(`Browser discovery: ${config}: ${elapsedMs} ms (exit ${result.exitCode}).`);
  if (result.exitCode !== 0)
    throw new TypeError(
      `Playwright discovery failed for ${config}:\n${result.stdout}\n${result.stderr}`,
    );
  const report: unknown = JSON.parse(result.stdout);
  if (!isRecord(report) || !Array.isArray(report["errors"]) || !Array.isArray(report["suites"]))
    throw new TypeError(`Invalid Playwright JSON discovery report for ${config}.`);
  if (report["errors"].length > 0)
    throw new TypeError(
      `Playwright discovery errors for ${config}: ${JSON.stringify(report["errors"])}`,
    );
  const tests = suiteTestCount(report["suites"]);
  if (tests === 0) throw new TypeError(`Playwright discovered no tests for ${config}.`);
  log(`Browser discovery: ${config}: ${tests} tests across all projects.`);
  return { config, tests };
};

export const checkBrowserDiscovery = ({
  repoRoot = REPO_ROOT,
  run = runCommand,
  cliPath,
  log = console.log,
}: BrowserDiscoveryOptions = {}) => {
  const coverage = run({
    command: [process.execPath, path.join(repoRoot, "scripts/check-browser-input-coverage.ts")],
    cwd: repoRoot,
  });
  if (coverage.exitCode !== 0)
    throw new TypeError(`Browser input coverage failed:\n${coverage.stdout}\n${coverage.stderr}`);
  if (coverage.stdout.trim() !== "") log(coverage.stdout.trim());
  const configs = trackedBrowserConfigs({ repoRoot, run });
  const results = configs.map((config) =>
    checkBrowserConfigDiscovery({ config, repoRoot, run, cliPath, log }),
  );
  log(`Browser discovery: ${results.length} tracked configs validated.`);
  return results;
};

if (import.meta.main) {
  try {
    if (process.argv.length !== 2) throw new TypeError("Browser discovery takes no arguments.");
    checkBrowserDiscovery();
  } catch (error) {
    process.exitCode = 1;
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  }
}
