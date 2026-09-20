import { expect, setDefaultTimeout, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const DEPCRUISE_BINARY = path.join(REPOSITORY_ROOT, "node_modules/.bin/depcruise");
const DEPCRUISE_CONFIG = path.join(REPOSITORY_ROOT, ".dependency-cruiser.cjs");
const KNOWN_VIOLATIONS_PATH = path.join(
  REPOSITORY_ROOT,
  ".dependency-cruiser-known-violations.json",
);
const NO_CIRCULAR_RULE_NAME = "no-circular";

type CycleModule = { name: string };

type KnownCycleViolation = {
  cycle: readonly CycleModule[];
  from: string;
  rule: { name: string };
  type: string;
};

type CruiseViolation = {
  cycle?: readonly CycleModule[];
  from: string;
  rule: { name: string };
};

setDefaultTimeout(60_000);

// A dependency-cruiser cycle violation is keyed by its member set, not by
// position: the same strongly connected component is reported once per
// module that starts the loop.
const sameCycle = (a: readonly CycleModule[], b: readonly CycleModule[]): boolean => {
  if (a.length !== b.length) return false;
  const bNames = new Set(b.map((module) => module.name));
  return a.every((module) => bNames.has(module.name));
};

const readKnownCycleViolations = (): KnownCycleViolation[] =>
  (JSON.parse(readFileSync(KNOWN_VIOLATIONS_PATH, "utf8")) as KnownCycleViolation[]).filter(
    (violation) => violation.type === "cycle" && violation.rule.name === NO_CIRCULAR_RULE_NAME,
  );

const readCurrentCycleViolations = (): CruiseViolation[] => {
  const result = Bun.spawnSync(
    [DEPCRUISE_BINARY, "--config", DEPCRUISE_CONFIG, "--output-type", "json", "packages"],
    {
      cwd: REPOSITORY_ROOT,
      env: { ...process.env, NODE_PATH: "node_modules" },
      stderr: "pipe",
      stdout: "pipe",
    },
  );
  const report = JSON.parse(result.stdout.toString()) as {
    summary: { violations: CruiseViolation[] };
  };
  return report.summary.violations.filter(
    (violation) => violation.rule.name === NO_CIRCULAR_RULE_NAME && violation.cycle,
  );
};

test("the circular-import allowlist may only shrink: every entry still reproduces", () => {
  const knownViolations = readKnownCycleViolations();
  if (knownViolations.length === 0) return;

  const currentViolations = readCurrentCycleViolations();
  const staleEntries = knownViolations.filter(
    (known) => !currentViolations.some((current) => sameCycle(known.cycle, current.cycle ?? [])),
  );

  expect(
    staleEntries.map((entry) => entry.from),
    "a cycle listed in .dependency-cruiser-known-violations.json no longer exists in the " +
      "code; remove its entry instead of leaving it allowlisted",
  ).toEqual([]);
});
