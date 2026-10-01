#!/usr/bin/env bun
// Consumer scenarios: run folio the way an external integrator does. Build and
// pack @stll/docx-core, @stll/folio-core, @stll/folio-agents and
// @stll/folio-cli, install the tarballs into a project OUTSIDE the monorepo,
// and run `test/consumer-scenarios/scenarios/*.test.ts` under Node's own test
// runner. Scenarios import only what the packed `exports` maps publish: the
// import guard below refuses anything else before a test runs, and Node's
// resolver refuses a subpath the map blocks.
//
// Flags:
//   --tarballs <dir>   reuse packed tarballs (<dir>/<package>/*.tgz) instead
//                      of building; `--pack-only <dir>` writes that layout
//   --keep             keep the staged consumer and print its path
//   --typecheck        also run `tsc --noEmit` over the scenarios in the stage
//   --only <pattern>   pass `--test-name-pattern` to node --test
//   --coverage-out <file>  where to write the coverage ledger summary
//                      (default: test-results/consumer-scenarios-coverage.json)
//   --feature-coverage-out <file>  operation × feature × selection report
//   [--] <files>       scenario files to run (default: all); Bun consumes a
//                      leading `--`, so bare file names are accepted too
//
// Coverage: every scenario process records which operation types met which
// stories, modes, target features and sessions (support/coverage.ts). The
// runner merges them, prints the table, writes the summary, and, when every
// scenario ran, fails if a cell test/consumer-scenarios/coverage-expectations.json
// requires was never applied or an unreachable cell was attempted.
//
// Environment: FOLIO_SCENARIO_SEED (fuzz seed; per commit under CI, fixed
// locally, `random` to explore; always printed), FOLIO_SCENARIO_FUZZ_RUNS / FOLIO_SCENARIO_FUZZ_STEPS
// (fuzz size; 12 runs of 10 steps by default), FOLIO_SCENARIO_COLLISION_RUNS
// (collision flows; 8 by default), FOLIO_ORACLE_GAPS=1 (print the operations
// the requested-outcome oracle could not model), FOLIO_SCENARIO_RELATIONS
// (metamorphic relations checked in the fuzz flows: `all` by default, `none`,
// or a comma-separated list) and FOLIO_SCENARIO_RELATIONS_DEPTH=full (check
// the sampled relations on every batch; for sweeps), FOLIO_SCENARIO_SAVE_SAMPLE
// (save this many initial generated flows as replayable DOCX artifacts) and
// FOLIO_SCENARIO_SAMPLE_DIR (artifact destination).
// FOLIO_SCENARIO_FEATURE_WEIGHTS=<report.json> enables bounded, seeded
// steering from a previous feature coverage report; unset keeps current draws.
// Flow files (test/consumer-scenarios/support/flow-file.ts):
// FOLIO_SCENARIO_FLOW (a flow file's JSON, or a path to one) replays that
// flow alone in flow-corpus.test.ts; FOLIO_SCENARIO_BUDGET_SECONDS turns on
// continuous-fuzz.test.ts, which reads and grows the corpus in
// FOLIO_SCENARIO_CORPUS_DIR and writes shrunk failures to
// FOLIO_SCENARIO_FAILURES_DIR and minimized new-state flows to
// FOLIO_SCENARIO_MINIMIZED_DIR (paths relative to where this runs).
// Exits non-zero on any failure, including missing required or hit unreachable coverage cells. Run via `bun run test:consumer-scenarios`.

import { panic } from "better-result";
import { $ } from "bun";
import { existsSync } from "node:fs";
import { cp, mkdir, mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { commitSeed } from "../test/commit-seed";
import {
  type Expectations,
  describePattern,
  formatLedger,
  hitUnreachableCells,
  mergeLedgers,
  mergeFeatureFiles,
  missingCells,
  summarize,
} from "../test/consumer-scenarios/support/coverage";
import {
  mergeFeatureCoverage,
  summarizeFeatureCoverage,
  type FeatureCoverage,
} from "../test/consumer-scenarios/support/feature-coverage";
import { buildAndPack, repoRoot } from "./packaged-consumer-lib";

const scenarioSrc = path.join(repoRoot, "test", "consumer-scenarios");

/** The packages a scenario may import, in dependency order. */
const PACKAGES = [
  { dir: "docx-core", name: "@stll/docx-core" },
  { dir: "core", name: "@stll/folio-core" },
  { dir: "agents", name: "@stll/folio-agents" },
  { dir: "cli", name: "@stll/folio-cli" },
] as const;

/** Third-party packages a consumer installs next to folio. */
const CONSUMER_DEPENDENCIES = ["prosemirror-state@^1.4.4", "prosemirror-model@^1.25.9"];
const TYPECHECK_DEPENDENCIES = ["typescript@^6", "@types/node@^22"];

type Args = {
  tarballs: string | null;
  packOnly: string | null;
  keep: boolean;
  typecheck: boolean;
  only: string | null;
  coverageOut: string;
  featureCoverageOut: string;
  files: string[];
};

const parseArgs = (argv: readonly string[]): Args => {
  const args: Args = {
    tarballs: null,
    packOnly: null,
    keep: false,
    typecheck: false,
    only: null,
    coverageOut: path.join(repoRoot, "test-results", "consumer-scenarios-coverage.json"),
    featureCoverageOut: path.join(repoRoot, "test-results", "consumer-scenarios-features.json"),
    files: [],
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const value = (): string => argv[++index] ?? panic(`consumer-scenarios: ${arg} needs a value`);
    if (arg === "--tarballs") args.tarballs = path.resolve(value());
    else if (arg === "--pack-only") args.packOnly = path.resolve(value());
    else if (arg === "--keep") args.keep = true;
    else if (arg === "--typecheck") args.typecheck = true;
    else if (arg === "--only") args.only = value();
    else if (arg === "--coverage-out") args.coverageOut = path.resolve(value());
    else if (arg === "--feature-coverage-out") args.featureCoverageOut = path.resolve(value());
    else if (arg === "--") args.files.push(...argv.slice(index + 1));
    else if (arg !== undefined && !arg.startsWith("-")) args.files.push(arg);
    else panic(`consumer-scenarios: unknown argument ${arg}`);
    if (arg === "--") break;
  }
  return args;
};

// ---------------------------------------------------------------------------
// Import guard
// ---------------------------------------------------------------------------

type ExportsMap = Record<string, unknown>;

/** `import … from "x"`, `export … from "x"`, `import "x"`, `import("x")`, `require("x")`. */
const IMPORT_SPECIFIER =
  /(?:\b(?:import|export)\s+(?:type\s+)?[\w*{}\s,$]*?\bfrom\s*|\bimport\s*\(\s*|\bimport\s+|\brequire\s*\(\s*)["'](?<specifier>[^"']+)["']/gu;

/** A conditional target's `import` (or `default`) file, a plain target as is. */
const importTarget = (target: unknown): string | null => {
  if (typeof target === "string") return target;
  if (typeof target !== "object" || target === null) return null;
  const conditions = target as Record<string, unknown>;
  return importTarget(conditions["import"] ?? conditions["default"] ?? null);
};

/**
 * The file `subpath` resolves to through an `exports` map, as Node resolves
 * it (an exact key first, else the pattern with the longest prefix before its
 * `*`), or `null` when the map does not export it.
 */
export const resolveExport = (exportsMap: ExportsMap, subpath: string): string | null => {
  if (Object.hasOwn(exportsMap, subpath)) {
    return importTarget(exportsMap[subpath]);
  }
  let best: { key: string; before: string; after: string } | null = null;
  for (const key of Object.keys(exportsMap)) {
    const star = key.indexOf("*");
    if (star === -1) continue;
    const before = key.slice(0, star);
    const after = key.slice(star + 1);
    if (
      subpath.startsWith(before) &&
      subpath.endsWith(after) &&
      subpath.length >= key.length - 1 &&
      (best === null || before.length > best.before.length)
    ) {
      best = { key, before, after };
    }
  }
  if (best === null) return null;
  const target = importTarget(exportsMap[best.key]);
  const matched = subpath.slice(best.before.length, subpath.length - best.after.length);
  return target === null ? null : target.replaceAll("*", matched);
};

const packageOf = (specifier: string): { name: string; subpath: string } => {
  const parts = specifier.split("/");
  const nameParts = specifier.startsWith("@") ? 2 : 1;
  const name = parts.slice(0, nameParts).join("/");
  const rest = parts.slice(nameParts).join("/");
  return { name, subpath: rest === "" ? "." : `./${rest}` };
};

type GuardContext = {
  /** Published `exports` and install directory of every folio package, by name. */
  exportsByPackage: ReadonlyMap<string, { exports: ExportsMap; dir: string }>;
  /** Other packages a scenario may import (by package name). */
  allowedPackages: ReadonlySet<string>;
};

/** Every import in `source` that reaches past a published entry point. */
export const findForbiddenImports = (source: string, context: GuardContext): string[] => {
  const problems: string[] = [];
  for (const match of source.matchAll(IMPORT_SPECIFIER)) {
    const specifier = match.groups?.["specifier"] ?? "";
    if (specifier.startsWith("node:")) continue;
    if (specifier.startsWith(".")) {
      // A scenario's own support modules, never a way around `exports`.
      if (/node_modules|\/dist\/|\/src\//u.test(specifier)) {
        problems.push(`${specifier}: reaches into an installed package's files`);
      }
      continue;
    }
    if (specifier.startsWith("/")) {
      problems.push(`${specifier}: absolute import`);
      continue;
    }
    const { name, subpath } = packageOf(specifier);
    const installed = context.exportsByPackage.get(name);
    if (installed) {
      const target = resolveExport(installed.exports, subpath);
      if (target === null || !existsSync(path.join(installed.dir, target))) {
        problems.push(`${specifier}: not a published entry point of ${name}`);
      }
      continue;
    }
    if (!context.allowedPackages.has(name)) {
      problems.push(`${specifier}: ${name} is not a consumer dependency`);
    }
  }
  return problems;
};

const listSourceFiles = async (dir: string): Promise<string[]> =>
  (await readdir(dir, { recursive: true }))
    .filter((file) => file.endsWith(".ts"))
    .map((file) => path.join(dir, file));

const guardImports = async (consumerDir: string): Promise<void> => {
  const exportsByPackage = new Map<string, { exports: ExportsMap; dir: string }>();
  for (const { name } of PACKAGES) {
    const dir = path.join(consumerDir, "node_modules", name);
    const manifest = JSON.parse(await readFile(path.join(dir, "package.json"), "utf8")) as {
      exports: ExportsMap;
    };
    exportsByPackage.set(name, { exports: manifest.exports, dir });
  }
  const context: GuardContext = {
    exportsByPackage,
    allowedPackages: new Set(
      CONSUMER_DEPENDENCIES.map((dependency) => dependency.slice(0, dependency.lastIndexOf("@"))),
    ),
  };

  // The guard guards itself: an internal path must be refused.
  const selfTest = findForbiddenImports(
    `import { x } from "@stll/folio-core/internal/anything";\nimport "@stll/folio-core/dist/server.js";`,
    context,
  );
  if (selfTest.length !== 2) {
    panic(`consumer-scenarios: the import guard let an internal path through (${selfTest})`);
  }

  const problems: string[] = [];
  for (const dir of ["scenarios", "support"]) {
    for (const file of await listSourceFiles(path.join(consumerDir, dir))) {
      for (const problem of findForbiddenImports(await readFile(file, "utf8"), context)) {
        problems.push(`${path.relative(consumerDir, file)}: ${problem}`);
      }
    }
  }
  if (problems.length > 0) {
    console.error(problems.map((problem) => `  ${problem}`).join("\n"));
    panic("consumer-scenarios: scenarios may import only published entry points");
  }
};

// ---------------------------------------------------------------------------
// Pack, stage, run
// ---------------------------------------------------------------------------

const packAll = async (destRoot: string): Promise<Map<string, string>> => {
  const tarballs = new Map<string, string>();
  for (const { dir, name } of PACKAGES) {
    const dest = path.join(destRoot, dir);
    await rm(dest, { recursive: true, force: true });
    await mkdir(dest, { recursive: true });
    tarballs.set(name, await buildAndPack(path.join(repoRoot, "packages", dir), dest));
  }
  return tarballs;
};

const findTarballs = async (root: string): Promise<Map<string, string>> => {
  const tarballs = new Map<string, string>();
  for (const { dir, name } of PACKAGES) {
    const files = (await readdir(path.join(root, dir))).filter((file) => file.endsWith(".tgz"));
    const [only] = files;
    if (files.length !== 1 || only === undefined) {
      panic(`consumer-scenarios: expected one tarball in ${path.join(root, dir)}`);
    }
    tarballs.set(name, path.join(root, dir, only));
  }
  return tarballs;
};

const stageConsumer = async (
  tarballs: ReadonlyMap<string, string>,
  typecheck: boolean,
): Promise<string> => {
  const consumerDir = await mkdtemp(path.join(tmpdir(), "folio-consumer-scenarios-"));
  console.log(`→ staging the consumer in ${consumerDir}`);
  for (const entry of ["scenarios", "support", "flows", "scenario-seeds.json", "tsconfig.json"]) {
    if (!existsSync(path.join(scenarioSrc, entry))) continue;
    await cp(path.join(scenarioSrc, entry), path.join(consumerDir, entry), { recursive: true });
  }
  const tarball = (name: string): string =>
    tarballs.get(name) ?? panic(`consumer-scenarios: no tarball for ${name}`);
  const manifest = {
    name: "folio-consumer-scenarios",
    version: "0.0.0",
    private: true,
    type: "module",
    // Pin every transitive folio package to its packed tarball; nothing here
    // may resolve from a registry or the workspace.
    overrides: Object.fromEntries(
      PACKAGES.filter(({ name }) => name !== "@stll/folio-cli").map(({ name }) => [
        name,
        tarball(name),
      ]),
    ),
  };
  await writeFile(path.join(consumerDir, "package.json"), `${JSON.stringify(manifest, null, 2)}\n`);
  console.log("→ installing the tarballs");
  const dependencies = [
    ...PACKAGES.map(({ name }) => tarball(name)),
    ...CONSUMER_DEPENDENCIES,
    ...(typecheck ? TYPECHECK_DEPENDENCIES : []),
  ];
  await $`bun add ${dependencies}`.cwd(consumerDir).quiet();
  return consumerDir;
};

/**
 * Merge the scenario processes' ledgers, print the table, write the summary,
 * and, when `enforce` (every scenario ran), check required and unreachable
 * cells. Returns the failure message, or null.
 */
const reportCoverage = async (
  coverageDir: string,
  out: string,
  featureOut: string,
  enforce: boolean,
): Promise<string | null> => {
  const expectations = JSON.parse(
    await readFile(path.join(scenarioSrc, "coverage-expectations.json"), "utf8"),
  ) as Expectations;
  const merged = mergeLedgers(coverageDir);
  console.log(`\n${formatLedger(merged)}`);
  await mkdir(path.dirname(out), { recursive: true });
  await writeFile(out, `${JSON.stringify(summarize(merged, expectations), null, 2)}\n`);
  console.log(`\n→ coverage ledger written to ${path.relative(repoRoot, out)}`);
  const conformanceInput = process.env["FOLIO_CONFORMANCE_FEATURE_COVERAGE_IN"];
  const conformance = conformanceInput
    ? (JSON.parse(await readFile(conformanceInput, "utf8")) as FeatureCoverage)
    : null;
  const features = mergeFeatureCoverage([
    mergeFeatureFiles(coverageDir),
    ...(conformance ? [conformance] : []),
  ]);
  const featureSummary = summarizeFeatureCoverage(features);
  await mkdir(path.dirname(featureOut), { recursive: true });
  await writeFile(featureOut, `${JSON.stringify(featureSummary, null, 2)}\n`);
  console.log(`→ feature coverage written to ${path.relative(repoRoot, featureOut)}`);
  console.log(`→ ${featureSummary.emptyCells.length} empty operation × feature × selection cells`);
  for (const cell of featureSummary.emptyCells.slice(0, 10)) console.log(`  no hits: ${cell}`);
  if (!enforce) {
    console.log("→ a subset ran; coverage expectations are not enforced");
    return null;
  }
  const missing = missingCells(merged, expectations);
  const unexpected = hitUnreachableCells(merged, expectations);
  const failures: string[] = [];
  if (missing.length === 0) {
    console.log(`✓ every required coverage cell was applied (${expectations.required.length})`);
  } else {
    console.error(missing.map((cell) => `  no hits: ${cell}`).join("\n"));
    failures.push(
      `✗ consumer-scenarios: ${missing.length} required coverage cell(s) had no applied operation (coverage-expectations.json).`,
    );
  }
  if (unexpected.length === 0) {
    console.log(
      `✓ every declared unreachable coverage cell had zero hits (${expectations.unreachable.length})`,
    );
  } else {
    console.error(
      unexpected
        .map(
          ({ cell, reason, applied, refused }) =>
            `  unexpected hits: ${describePattern(cell)} (${applied} applied, ${refused} refused; ${reason})`,
        )
        .join("\n"),
    );
    failures.push(
      `✗ consumer-scenarios: ${unexpected.length} declared unreachable coverage cell(s) had hits (coverage-expectations.json).`,
    );
  }
  return failures.length === 0 ? null : failures.join("\n");
};

const args = parseArgs(process.argv.slice(2));

if (args.packOnly !== null) {
  const tarballs = await packAll(args.packOnly);
  console.log([...tarballs.values()].join("\n"));
  process.exit(0);
}

// Under CI the fuzz base derives from the commit (test/commit-seed.ts): every
// commit searches new flows, yet a rerun of the same commit replays exactly,
// so a red run stays red. Locally the fixed default replays the same flows;
// `FOLIO_SCENARIO_SEED=random` explores. The seed is printed either way and
// replays with the same value. Pinned seeds in the scenarios run regardless.
const DEFAULT_SEED = "20260926";
/** A non-negative base with room for `seed + run` to stay a safe integer. */
const ciSeed = (): string | null => {
  const ci = process.env["CI"];
  if (ci === undefined || ci === "" || ci === "false" || ci === "0") return null;
  const derived = commitSeed("consumer-scenarios fuzz");
  return derived === undefined ? null : String(derived >>> 1);
};
const requestedSeed = process.env["FOLIO_SCENARIO_SEED"] ?? ciSeed() ?? DEFAULT_SEED;
const seed =
  requestedSeed === "random" ? String(Math.floor(Math.random() * 2 ** 31)) : requestedSeed;

/**
 * The flow-file settings as the staged scenarios need them: directories made
 * absolute (the scenarios run elsewhere), a flow file path read into JSON.
 */
const flowEnvironment = async (): Promise<Record<string, string>> => {
  const environment: Record<string, string> = {};
  for (const name of [
    "FOLIO_SCENARIO_CORPUS_DIR",
    "FOLIO_SCENARIO_FAILURES_DIR",
    "FOLIO_SCENARIO_MINIMIZED_DIR",
  ]) {
    const value = process.env[name];
    if (value !== undefined && value !== "") environment[name] = path.resolve(value);
  }
  // Fingerprints with an open issue: continuous-fuzz.test.ts records them without failing.
  const known = JSON.parse(
    await readFile(path.join(repoRoot, "test", "known-failure-fingerprints.json"), "utf8"),
  ) as { known: { fingerprint: string }[] };
  environment["FOLIO_SCENARIO_KNOWN_FINGERPRINTS"] = known.known
    .map(({ fingerprint }) => fingerprint)
    .join(",");
  const flow = process.env["FOLIO_SCENARIO_FLOW"]?.trim();
  if (flow !== undefined && flow !== "") {
    environment["FOLIO_SCENARIO_FLOW"] = flow.startsWith("{")
      ? flow
      : await readFile(path.resolve(flow), "utf8");
  }
  return environment;
};

let failure: string | null = null;
let packRoot = "";
let consumerDir = "";
try {
  let tarballs: Map<string, string>;
  if (args.tarballs === null) {
    packRoot = await mkdtemp(path.join(tmpdir(), "folio-consumer-scenarios-pack-"));
    tarballs = await packAll(packRoot);
  } else {
    tarballs = await findTarballs(args.tarballs);
  }
  consumerDir = await stageConsumer(tarballs, args.typecheck);
  await guardImports(consumerDir);

  if (args.typecheck) {
    console.log("→ typechecking the scenarios against the published declarations");
    const typecheck = await $`node node_modules/typescript/bin/tsc --noEmit -p tsconfig.json`
      .cwd(consumerDir)
      .nothrow();
    if (typecheck.exitCode !== 0) {
      console.error(typecheck.stdout.toString() || typecheck.stderr.toString());
      failure = "✗ consumer-scenarios: the scenarios do not typecheck against the tarballs.";
    }
  }

  if (failure === null) {
    const files =
      args.files.length > 0
        ? args.files.map((file) => path.join("scenarios", path.basename(file)))
        : (await readdir(path.join(consumerDir, "scenarios")))
            .filter((file) => file.endsWith(".test.ts"))
            .sort()
            .map((file) => path.join("scenarios", file));
    const nameFilter = args.only === null ? [] : ["--test-name-pattern", args.only];
    console.log(`→ node --test over ${files.length} scenario files (seed ${seed})`);
    const coverageDir = path.join(consumerDir, "coverage");
    const run = await $`node --test --test-reporter=spec ${nameFilter} ${files}`
      .cwd(consumerDir)
      .env({
        ...process.env,
        ...(await flowEnvironment()),
        FOLIO_SCENARIO_SEED: seed,
        FOLIO_SCENARIO_COVERAGE_DIR: coverageDir,
        FOLIO_SCENARIO_SAMPLE_DIR:
          process.env["FOLIO_SCENARIO_SAMPLE_DIR"] ??
          path.join(repoRoot, "test-results", "consumer-scenario-samples"),
        ...(process.env["FOLIO_SCENARIO_FEATURE_WEIGHTS"]
          ? {
              FOLIO_SCENARIO_FEATURE_WEIGHTS: path.resolve(
                process.env["FOLIO_SCENARIO_FEATURE_WEIGHTS"],
              ),
            }
          : {}),
      })
      .nothrow();
    if (run.exitCode !== 0) {
      failure = `✗ consumer-scenarios: scenarios failed (FOLIO_SCENARIO_SEED=${seed} reproduces the fuzz runs).`;
    }
    const coverageFailure = await reportCoverage(
      coverageDir,
      args.coverageOut,
      args.featureCoverageOut,
      args.files.length === 0 && args.only === null,
    );
    failure ??= coverageFailure;
  }
} finally {
  if (packRoot !== "") {
    await rm(packRoot, { recursive: true, force: true });
  }
  if (consumerDir !== "") {
    if (args.keep) {
      console.log(`→ kept the staged consumer at ${consumerDir}`);
    } else {
      await rm(consumerDir, { recursive: true, force: true });
    }
  }
}

if (failure !== null) {
  console.error(failure);
  process.exit(1);
}
console.log(
  `\n✓ consumer-scenarios: every scenario passed against the packed tarballs (seed ${seed}).`,
);
