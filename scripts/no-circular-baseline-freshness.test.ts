import { expect, setDefaultTimeout, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const DEPCRUISE_BINARY = path.join(REPOSITORY_ROOT, "node_modules/.bin/depcruise");
const DEPCRUISE_CONFIG = path.join(REPOSITORY_ROOT, ".dependency-cruiser.cjs");
const TYPESCRIPT_PRELOAD = path.join(
  REPOSITORY_ROOT,
  "scripts/lib/depcruise-typescript-preload.mjs",
);
const KNOWN_VIOLATIONS_PATH = path.join(
  REPOSITORY_ROOT,
  ".dependency-cruiser-known-violations.json",
);
const TYPE_ONLY_DEPENDENCY_TYPE = "type-only";

type CycleModule = { name: string };

type KnownViolation = {
  cycle?: readonly CycleModule[];
  from: string;
  rule: { name: string };
  to?: string;
  type: string;
};

type CruiseViolation = {
  cycle?: readonly CycleModule[];
  from: string;
  rule: { name: string };
  to?: string;
};

type CruiseRun = {
  dependencyTypesSeen: ReadonlySet<string>;
  violations: readonly CruiseViolation[];
};

setDefaultTimeout(60_000);

// A dependency-cruiser cycle violation is keyed by its member set, not by
// position: the same strongly connected component is reported once per
// module that starts the loop. A non-cycle (plain "dependency") violation is
// keyed by its edge, matching dependency-cruiser's own known-violations
// softening logic (src/analyze/summarize/is-same-violation.mjs).
const isSameViolation = (known: KnownViolation, current: CruiseViolation): boolean => {
  if (known.rule.name !== current.rule.name) return false;
  if (known.cycle && current.cycle) {
    if (known.cycle.length !== current.cycle.length) return false;
    const currentNames = new Set(current.cycle.map((module) => module.name));
    return known.cycle.every((module) => currentNames.has(module.name));
  }
  return known.from === current.from && known.to === current.to;
};

const readKnownViolations = (): KnownViolation[] =>
  JSON.parse(readFileSync(KNOWN_VIOLATIONS_PATH, "utf8")) as KnownViolation[];

const runCruise = (): CruiseRun => {
  const result = Bun.spawnSync(
    [DEPCRUISE_BINARY, "--config", DEPCRUISE_CONFIG, "--output-type", "json", "packages"],
    {
      cwd: REPOSITORY_ROOT,
      env: {
        ...process.env,
        NODE_OPTIONS: `--import=${TYPESCRIPT_PRELOAD}`,
        NODE_PATH: "node_modules",
      },
      stderr: "pipe",
      stdout: "pipe",
    },
  );
  const report = JSON.parse(result.stdout.toString()) as {
    modules: ReadonlyArray<{
      dependencies: ReadonlyArray<{ dependencyTypes?: readonly string[] }>;
    }>;
    summary: { violations: readonly CruiseViolation[] };
  };
  const dependencyTypesSeen = new Set<string>();
  for (const module of report.modules) {
    for (const dependency of module.dependencies) {
      for (const dependencyType of dependency.dependencyTypes ?? []) {
        dependencyTypesSeen.add(dependencyType);
      }
    }
  }
  return { dependencyTypesSeen, violations: report.summary.violations };
};

test("the tsc extractor is active: it tags at least one edge type-only", () => {
  // no-circular's type-only exclusion (see .dependency-cruiser.cjs) is only
  // as good as this tag being produced at all. dependency-cruiser silently
  // falls back to an extractor that never produces it when its tsc-backed
  // one can't load typescript (the exact failure this preload fixes); without
  // this assertion that fallback would quietly widen every rule that keys off
  // "type-only" back to matching every import, re-inflating the baseline
  // instead of failing loudly.
  const { dependencyTypesSeen } = runCruise();
  expect(dependencyTypesSeen.has(TYPE_ONLY_DEPENDENCY_TYPE)).toBe(true);
});

test("the known-violations allowlist may only shrink: every entry still reproduces", () => {
  const knownViolations = readKnownViolations();
  if (knownViolations.length === 0) return;

  const { violations: currentViolations } = runCruise();
  const staleEntries = knownViolations.filter(
    (known) => !currentViolations.some((current) => isSameViolation(known, current)),
  );

  expect(
    staleEntries.map((entry) => `${entry.rule.name}: ${entry.from} -> ${entry.to ?? "(cycle)"}`),
    "an entry in .dependency-cruiser-known-violations.json no longer reproduces; remove it " +
      "instead of leaving it allowlisted",
  ).toEqual([]);
});
