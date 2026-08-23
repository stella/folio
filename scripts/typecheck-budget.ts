#!/usr/bin/env bun

import { readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const BASELINE_PATH = path.join(import.meta.dir, "typecheck-budget.json");
const BASELINE_RELATIVE_PATH = "scripts/typecheck-budget.json";
const NATIVE_TSC = path.join(REPO_ROOT, "node_modules/@typescript/native/bin/tsc");
const CLASSIC_TSC = path.join(REPO_ROOT, "node_modules/typescript/bin/tsc");

const PROJECTS = [
  { id: "docx-core", config: "packages/docx-core/tsconfig.build.json", compiler: "native" },
  { id: "core", config: "packages/core/tsconfig.build.json", compiler: "native" },
  { id: "react", config: "packages/react/tsconfig.build.json", compiler: "native" },
  { id: "agents", config: "packages/agents/tsconfig.build.json", compiler: "native" },
  { id: "vue", config: "packages/vue/tsconfig.build.json", compiler: "vue" },
  { id: "nuxt", config: "packages/nuxt/tsconfig.json", compiler: "classic" },
  { id: "playground", config: "packages/playground/tsconfig.json", compiler: "native" },
  { id: "playground-vue", config: "packages/playground-vue/tsconfig.json", compiler: "vue" },
  { id: "parity", config: "parity/tsconfig.check.json", compiler: "classic" },
] as const;

const GATED_FIELDS = ["types", "instantiations"] as const;
const HEADROOM = 1.05;
const HEADROOM_FLOOR = {
  types: 1_000,
  instantiations: 5_000,
} satisfies Counters;
const RATCHET_THRESHOLD = 0.95;
const WRITE_COMMAND = "bun run typecheck:budget:write";

type ProjectId = (typeof PROJECTS)[number]["id"];
type GatedField = (typeof GATED_FIELDS)[number];
type Counters = Record<GatedField, number>;
type Baseline = Record<ProjectId, Counters>;
type Measurement = {
  id: ProjectId;
  counters: Counters;
  context: string;
};

const readDiagnostic = (output: string, label: string): number | null => {
  const match = new RegExp(`^${label}:\\s+([\\d.]+)`, "mu").exec(output);
  const value = match?.at(1);
  return value === undefined ? null : Number(value);
};

const isCounter = (value: unknown): value is number =>
  typeof value === "number" && Number.isFinite(value) && value >= 0;

type KeyDifferences = { missing: string[]; extra: string[] };

const keyDifferences = (
  actual: readonly string[],
  expected: readonly string[],
): KeyDifferences => ({
  missing: expected.filter((key) => !actual.includes(key)),
  extra: actual.filter((key) => !expected.includes(key)),
});

const assertExactKeys = (
  value: unknown,
  expected: readonly string[],
  label: string,
): asserts value is Record<string, unknown> => {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    throw new Error(`Invalid ${label}: expected an object`);
  }
  const { missing, extra } = keyDifferences(Object.keys(value), expected);
  if (missing.length > 0 || extra.length > 0) {
    throw new Error(
      `Invalid ${label} keys; missing: ${missing.join(", ") || "none"}; ` +
        `extra: ${extra.join(", ") || "none"}`,
    );
  }
};

const rejectsExactKeys = (value: unknown, expected: readonly string[]): boolean => {
  try {
    assertExactKeys(value, expected, "self-test baseline");
    return false;
  } catch {
    return true;
  }
};

const parseCounters = (output: string, project: string): Counters => {
  const types = readDiagnostic(output, "Types");
  const instantiations = readDiagnostic(output, "Instantiations");
  if (types === null || instantiations === null) {
    throw new Error(`Missing Types/Instantiations in diagnostics for ${project}:\n${output}`);
  }
  return { types, instantiations };
};

const compilerExecutable = (
  compiler: (typeof PROJECTS)[number]["compiler"],
  config: string,
): string => {
  if (compiler === "native") {
    return NATIVE_TSC;
  }
  if (compiler === "vue") {
    return path.join(REPO_ROOT, path.dirname(config), "node_modules/vue-tsc/bin/vue-tsc.js");
  }
  return CLASSIC_TSC;
};

const measureProject = ({ id, config, compiler }: (typeof PROJECTS)[number]): Measurement => {
  const executable = compilerExecutable(compiler, config);
  const flags = ["-p", config, "--noEmit", "--extendedDiagnostics"];
  if (compiler === "native") {
    flags.push("--singleThreaded");
  }
  flags.push("--incremental", "false");
  // vue-tsc installs a Node launcher that patches TypeScript's module loading;
  // invoking that launcher through Bun skips its Vue SFC integration.
  const command =
    compiler === "vue" ? [executable, ...flags] : [process.execPath, executable, ...flags];
  const spawned = Bun.spawnSync(command, {
    cwd: REPO_ROOT,
    stdout: "pipe",
    stderr: "pipe",
  });
  const output = spawned.stdout.toString();
  if (spawned.exitCode !== 0) {
    throw new Error(
      `Typecheck failed for ${config} (exit ${spawned.exitCode}).\n${output}${spawned.stderr.toString()}`,
    );
  }
  const files = readDiagnostic(output, "Files");
  const checkTime = readDiagnostic(output, "Check time");
  const memoryKiB = readDiagnostic(output, "Memory used");
  const memory = memoryKiB === null ? "?" : `${Math.round(memoryKiB / 1024)} MiB`;
  return {
    id,
    counters: parseCounters(output, config),
    context: `${files ?? "?"} files, ${checkTime ?? "?"}s check, ${memory}`,
  };
};

const measureAll = (): Measurement[] =>
  PROJECTS.map((project) => {
    console.log(`typecheck-budget: measuring ${project.id}`);
    return measureProject(project);
  });

const readBaselineEntry = (parsed: Record<string, unknown>, id: ProjectId): Counters => {
  const value = parsed[id];
  if (
    typeof value !== "object" ||
    value === null ||
    !("types" in value) ||
    !isCounter(value.types) ||
    !("instantiations" in value) ||
    !isCounter(value.instantiations)
  ) {
    throw new Error(`Invalid ${id} entry in ${BASELINE_RELATIVE_PATH}`);
  }
  return { types: value.types, instantiations: value.instantiations };
};

const readBaseline = (): Baseline => {
  const parsed: unknown = JSON.parse(readFileSync(BASELINE_PATH, "utf8"));
  assertExactKeys(
    parsed,
    PROJECTS.map(({ id }) => id),
    BASELINE_RELATIVE_PATH,
  );
  return {
    "docx-core": readBaselineEntry(parsed, "docx-core"),
    core: readBaselineEntry(parsed, "core"),
    react: readBaselineEntry(parsed, "react"),
    agents: readBaselineEntry(parsed, "agents"),
    vue: readBaselineEntry(parsed, "vue"),
    nuxt: readBaselineEntry(parsed, "nuxt"),
    playground: readBaselineEntry(parsed, "playground"),
    "playground-vue": readBaselineEntry(parsed, "playground-vue"),
    parity: readBaselineEntry(parsed, "parity"),
  };
};

const allowedMaximum = (field: GatedField, baseline: number): number =>
  Math.max(baseline * HEADROOM, baseline + HEADROOM_FLOOR[field]);

const exceedsMaximum = (field: GatedField, current: number, baseline: number): boolean =>
  current > allowedMaximum(field, baseline);

const ratchetAvailable = (current: number, baseline: number): boolean =>
  current < baseline * RATCHET_THRESHOLD;

const percentage = (current: number, baseline: number): string => {
  const delta = ((current - baseline) / baseline) * 100;
  return `${delta >= 0 ? "+" : ""}${delta.toFixed(1)}%`;
};

const formatted = (value: number): string => value.toLocaleString("en-US");

const writeBaseline = (measurements: readonly Measurement[]): void => {
  const baseline = Object.fromEntries(
    PROJECTS.map(({ id }) => {
      const measurement = measurements.find((entry) => entry.id === id);
      if (measurement === undefined) {
        throw new Error(`Missing measurement for ${id}`);
      }
      return [id, measurement.counters];
    }),
  );
  writeFileSync(BASELINE_PATH, `${JSON.stringify(baseline, null, 2)}\n`);
};

const report = (measurements: readonly Measurement[]): void => {
  for (const { id, counters, context } of measurements) {
    console.log(
      `typecheck-budget: ${id} types ${formatted(counters.types)}, ` +
        `instantiations ${formatted(counters.instantiations)} (${context})`,
    );
  }
};

const check = (measurements: readonly Measurement[]): number => {
  const baseline = readBaseline();
  const failures: string[] = [];
  const improvements: string[] = [];
  for (const measurement of measurements) {
    const expected = baseline[measurement.id];
    for (const field of GATED_FIELDS) {
      const current = measurement.counters[field];
      const previous = expected[field];
      if (exceedsMaximum(field, current, previous)) {
        failures.push(
          `${measurement.id}.${field}: ${formatted(previous)} -> ${formatted(current)} ` +
            `(${percentage(current, previous)})`,
        );
      } else if (ratchetAvailable(current, previous)) {
        improvements.push(
          `${measurement.id}.${field}: ${formatted(previous)} -> ${formatted(current)}`,
        );
      }
    }
  }
  for (const improvement of improvements) {
    console.log(`typecheck-budget: ratchet available for ${improvement}; run ${WRITE_COMMAND}`);
  }
  if (failures.length === 0) {
    console.log("typecheck-budget: all compiler-work budgets are within limits");
    return 0;
  }
  console.error("typecheck-budget: compiler workload exceeded its maximum:");
  for (const failure of failures) {
    console.error(`  ${failure}`);
  }
  console.error(
    `Fix the inference expansion. If the increase is intentional, run ${WRITE_COMMAND} and justify the new baseline.`,
  );
  return 1;
};

const selfTest = (): number => {
  const sample = [
    "Files:             790",
    "Types:          105410",
    "Instantiations:  74508",
    "Memory used:   273905K",
    "Check time:     0.572s",
  ].join("\n");
  const parsed = parseCounters(sample, "self-test");
  if (parsed.types !== 105_410 || parsed.instantiations !== 74_508) {
    console.error("typecheck-budget: diagnostics parser self-test failed");
    return 1;
  }
  if (
    allowedMaximum("instantiations", 1_000_000) !== 1_050_000 ||
    allowedMaximum("types", 10_000) !== 11_000 ||
    exceedsMaximum("types", 11_000, 10_000) ||
    !exceedsMaximum("types", 11_001, 10_000) ||
    ratchetAvailable(9_500, 10_000) ||
    !ratchetAvailable(9_499, 10_000)
  ) {
    console.error("typecheck-budget: comparison self-test failed");
    return 1;
  }
  if (
    rejectsExactKeys({ core: {} }, ["core"]) ||
    !rejectsExactKeys({}, ["core"]) ||
    !rejectsExactKeys({ core: {}, stale: {} }, ["core"])
  ) {
    console.error("typecheck-budget: baseline-key self-test failed");
    return 1;
  }

  const expectedWorkspaceProjects = PROJECTS.filter(({ config }) =>
    config.startsWith("packages/"),
  ).map(({ id }) => id);
  const actualWorkspaceProjects: string[] = [];
  const packagesRoot = path.join(REPO_ROOT, "packages");
  for (const entry of readdirSync(packagesRoot, { withFileTypes: true })) {
    if (!entry.isDirectory()) {
      continue;
    }
    const packageJsonPath = path.join(packagesRoot, entry.name, "package.json");
    const packageJson: unknown = JSON.parse(readFileSync(packageJsonPath, "utf8"));
    if (typeof packageJson !== "object" || packageJson === null || !("scripts" in packageJson)) {
      continue;
    }
    const { scripts } = packageJson;
    if (
      typeof scripts === "object" &&
      scripts !== null &&
      "typecheck" in scripts &&
      typeof scripts.typecheck === "string"
    ) {
      actualWorkspaceProjects.push(entry.name);
    }
  }
  const workspaceDiff = keyDifferences(actualWorkspaceProjects, expectedWorkspaceProjects);
  if (workspaceDiff.missing.length > 0 || workspaceDiff.extra.length > 0) {
    console.error(
      "typecheck-budget: workspace coverage self-test failed; " +
        `missing: ${workspaceDiff.missing.join(", ") || "none"}; ` +
        `extra: ${workspaceDiff.extra.join(", ") || "none"}`,
    );
    return 1;
  }
  console.log("typecheck-budget: self-test passed");
  return 0;
};

const mode = process.argv.at(2) ?? "--report";
if (process.argv.length > 3) {
  console.error("Usage: typecheck-budget.ts [--check|--report|--self-test|--write]");
  process.exit(2);
}
if (mode === "--self-test") {
  process.exit(selfTest());
}
if (!["--check", "--report", "--write"].includes(mode)) {
  console.error("Usage: typecheck-budget.ts [--check|--report|--self-test|--write]");
  process.exit(2);
}

try {
  const measurements = measureAll();
  report(measurements);
  if (mode === "--write") {
    writeBaseline(measurements);
    console.log(`typecheck-budget: wrote ${BASELINE_RELATIVE_PATH}`);
    process.exit(0);
  }
  process.exit(mode === "--check" ? check(measurements) : 0);
} catch (error) {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
}
