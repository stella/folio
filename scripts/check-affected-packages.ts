#!/usr/bin/env bun
import { affectedPackagePlan, changedWorkingTreePaths, publishedPackagePaths } from "./ci-api-plan";

const usage = "Usage: bun run check:affected-packages --base <rev>";

const parseBase = (args: readonly string[]) => {
  if (args.length !== 2 || args.at(0) !== "--base" || !args.at(1)) throw new TypeError(usage);
  return args[1];
};

type RunOptions = { env?: Record<string, string> };
const run = (command: readonly string[], options: RunOptions = {}) => {
  process.stdout.write(`Running: ${command.join(" ")}\n`);
  const result = Bun.spawnSync(command, {
    cwd: process.cwd(),
    stdin: "inherit",
    stdout: "inherit",
    stderr: "inherit",
    env: { ...process.env, ...options.env },
  });
  return result.exitCode;
};

const main = () => {
  const base = parseBase(process.argv.slice(2));
  const root = Bun.spawnSync(["git", "rev-parse", "--show-toplevel"], {
    stdout: "pipe",
    stderr: "pipe",
  });
  if (root.exitCode !== 0)
    throw new TypeError(`Cannot locate repository: ${root.stderr.toString().trim()}`);
  const repoRoot = root.stdout.toString().trim();
  process.chdir(repoRoot);
  const changedPaths = changedWorkingTreePaths({ repoRoot, base });
  const plan = affectedPackagePlan({
    changedPaths,
    publishedPackagePaths: publishedPackagePaths(repoRoot),
  });
  process.stdout.write(
    `Affected-package plan from ${base}: ${changedPaths.length} changed file(s).\n`,
  );

  if (plan.apiRequired) {
    for (const command of [
      ["bun", "run", "build"],
      ["bun", "run", "check:build-reproducibility"],
      ["bun", "run", "api:check"],
    ]) {
      const exitCode = run(command);
      if (exitCode === 0) continue;
      if (command.at(-1) === "api:check")
        process.stderr.write("run `bun run api:update` and commit api-reports\n");
      return exitCode;
    }
  } else {
    process.stdout.write("Skipped API checks: no published package API inputs changed.\n");
  }

  if (plan.packagedConsumerRequired) {
    process.stdout.write(
      "Running the reduced consumer suite (typecheck plus compare.test.ts); the full suite takes far above a few minutes.\n",
    );
    const exitCode = run([
      "bun",
      "scripts/consumer-scenarios.ts",
      "--typecheck",
      "compare.test.ts",
    ]);
    if (exitCode !== 0) return exitCode;
  } else {
    process.stdout.write(
      "Skipped consumer scenarios: no packed package or scenario inputs changed.\n",
    );
  }
  return 0;
};

if (import.meta.main) {
  try {
    process.exitCode = main();
  } catch (error) {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  }
}
