#!/usr/bin/env bun

import path from "node:path";

// Raise this after checking the real cruise summary when sustained graph growth
// leaves substantial headroom; keep room for ordinary module consolidation.
export const MINIMUM_CRUISED_MODULES = 2400;

export const dependencyCruiseCoverageIssue = (moduleCount: unknown): string | null => {
  if (typeof moduleCount !== "number" || !Number.isSafeInteger(moduleCount) || moduleCount < 0) {
    return "dependency-cruiser did not report a valid cruised module count";
  }
  if (moduleCount < MINIMUM_CRUISED_MODULES) {
    return `dependency-cruiser cruised ${moduleCount} modules; expected at least ${MINIMUM_CRUISED_MODULES}. Check source discovery and the TypeScript extractor.`;
  }
  return null;
};

export const dependencyCruiseOutputIssue = (output: string): string | null => {
  // Both successful and violation summaries use these fields in the err reporter.
  const summaries = [...output.matchAll(/\b(\d+) modules, \d+ dependencies cruised\b/gu)];
  if (summaries.length !== 1) {
    return "dependency-cruiser did not report exactly one cruise summary";
  }
  return dependencyCruiseCoverageIssue(Number(summaries.at(0)?.at(1)));
};

const main = async () => {
  const root = path.resolve(import.meta.dirname, "..");
  const child = Bun.spawn(
    [
      path.join(root, "node_modules/.bin/depcruise"),
      "--config",
      ".dependency-cruiser.cjs",
      "--ignore-known",
      ".dependency-cruiser-known-violations.json",
      "--output-type",
      "err",
      "packages",
    ],
    {
      cwd: root,
      env: {
        ...process.env,
        NODE_PATH: "node_modules",
        NODE_OPTIONS: `--import=${path.join(root, "scripts/lib/depcruise-typescript-preload.mjs")}`,
      },
      stdout: "pipe",
      stderr: "inherit",
    },
  );
  const [output, exitCode] = await Promise.all([new Response(child.stdout).text(), child.exited]);
  process.stdout.write(output);
  const issue = dependencyCruiseOutputIssue(output);
  if (issue !== null) process.stderr.write(`${issue}\n`);
  process.exitCode = exitCode || (issue === null ? 0 : 1);
};

if (import.meta.main) await main();
