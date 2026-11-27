#!/usr/bin/env bun
/**
 * The source guards CI runs on every pull request, run locally over the files
 * a change touches.
 *
 * It compares the branch, with its uncommitted changes, against the merge base
 * with `origin/main`, picks the guards whose inputs the change touched, and
 * runs them a few at a time. The two shrink-only baselines check only the
 * touched files' entries, so a run takes seconds rather than the minutes the
 * whole-tree scans take. Lint, typecheck, tests and builds stay with CI.
 *
 *   bun scripts/check-pr-guards.ts               run the guards for the change
 *   bun scripts/check-pr-guards.ts --list        name them without running
 *   bun scripts/check-pr-guards.ts --base <ref>  compare against another ref
 */

import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");

/** How many guards run at once. */
const CONCURRENCY = 3;

export type PrGuard = {
  readonly id: string;
  readonly command: readonly string[];
  readonly fixHint: string;
};

const isPackageSource = (file: string): boolean =>
  /^packages\/[^/]+\/src\//u.test(file);

const isManifest = (file: string): boolean =>
  file === "bun.lock" ||
  file === "package.json" ||
  /^(?:packages|plugins|benchmarks|playground)[^/]*\/(?:[^/]+\/)*package\.json$/u.test(
    file,
  );

const isCatalogSource = (file: string): boolean =>
  file.startsWith("packages/core/src/i18n/") ||
  (file.startsWith("packages/") && /\.(?:tsx|vue)$/u.test(file));

const isAdapterSource = (file: string): boolean =>
  file.startsWith("packages/react/") || file.startsWith("packages/vue/");

/** The scripts that check a generated file, from the `generate:*:check` names. */
export const generatedChecks = (
  scripts: Readonly<Record<string, string>>,
): readonly string[] =>
  Object.keys(scripts)
    .filter((name) => /^generate:[a-z0-9-]+:check$/u.test(name))
    .toSorted();

/**
 * The Unicode tables read sources `specifications:fetch` downloads, so their
 * check runs only when its own inputs change, after the fetch.
 */
const JOINING_TYPES_CHECK = "generate:joining-types:check";
const JOINING_TYPES_INPUTS = [
  "scripts/generate-joining-types.ts",
  "packages/core/src/utils/joiningTypes.gen.ts",
  "specifications/sources.json",
];

const hasScript = (
  scripts: Readonly<Record<string, string>>,
  name: string,
): boolean => Object.hasOwn(scripts, name);

const run = (name: string): readonly string[] => ["bun", "run", name];

/** The guards relevant to `changedFiles`, in the order CI lists them. */
export const selectPrGuards = (
  changedFiles: readonly string[],
  scripts: Readonly<Record<string, string>>,
): readonly PrGuard[] => {
  const touches = (predicate: (file: string) => boolean): boolean =>
    changedFiles.some(predicate);
  const guards: PrGuard[] = [];

  if (touches(isManifest)) {
    for (const script of [
      "check:lockfile-versions",
      "check:lockfile-overrides",
      "check:lockfile-trust",
      "check:dedupe",
    ])
      guards.push({
        id: script,
        command: run(script),
        fixHint: "Run `bun install` and commit bun.lock.",
      });
  }

  const sources = changedFiles.filter(isPackageSource);
  const generatorInputs = touches(
    (file) =>
      isPackageSource(file) ||
      file.startsWith("scripts/generate-") ||
      file.startsWith("specifications/"),
  );

  if (sources.length > 0) {
    guards.push(
      {
        id: "check:normalization-warnings",
        command: run("check:normalization-warnings"),
        fixHint:
          "Name a warning code in the normaliser, or state why it loses nothing.",
      },
      {
        id: "check:canonical-cutover",
        command: run("check:canonical-cutover"),
        fixHint:
          "Run `bun scripts/check-canonical-cutover.ts --write` and commit the ledger.",
      },
      {
        id: "check:strict-name-tolerance",
        command: run("check:strict-name-tolerance"),
        fixHint: "Read both spellings through the Strict names table.",
      },
      {
        id: "check:reserved-value-coverage",
        command: run("check:reserved-value-coverage"),
        fixHint: "Register or exclude the new reserved-value slot.",
      },
      {
        id: "check:reserved-values",
        command: [
          "bun",
          "scripts/reserved-value-baseline.ts",
          "--check",
          "--files",
          ...sources,
        ],
        fixHint:
          "Read the value through its registry reader, or run `bun scripts/reserved-value-baseline.ts --write-baseline` if comparisons only went away.",
      },
      {
        id: "check:identity-attributes",
        command: [
          "bun",
          "scripts/identity-attribute-baseline.ts",
          "--check",
          "--files",
          ...sources,
        ],
        fixHint:
          "Resolve the attribute by namespace URI, or run `bun scripts/identity-attribute-baseline.ts --write-baseline` if reads only went away.",
      },
    );
  }

  if (generatorInputs) {
    for (const script of generatedChecks(scripts)) {
      if (script === JOINING_TYPES_CHECK) continue;
      guards.push({
        id: script,
        command: run(script),
        fixHint: `Run \`bun run ${script.replace(/:check$/u, "")}\` and commit the generated file.`,
      });
    }
  }

  if (
    hasScript(scripts, JOINING_TYPES_CHECK) &&
    touches((file) => JOINING_TYPES_INPUTS.includes(file))
  ) {
    guards.push({
      id: JOINING_TYPES_CHECK,
      command: [
        "sh",
        "-c",
        "bun run specifications:fetch --profile shaping && bun run generate:joining-types:check",
      ],
      fixHint: "Run `bun run generate:joining-types` and commit the table.",
    });
  }

  if (touches(isCatalogSource)) {
    guards.push({
      id: "check:i18n",
      command: run("check:i18n"),
      fixHint: "Run `bun run i18n:sync` and add the missing messages.",
    });
  }

  if (touches(isAdapterSource)) {
    for (const id of ["check:parity-contract", "check:export-parity"])
      guards.push({
        id,
        command: run(id),
        fixHint: "Give the React and Vue adapters the same props and exports.",
      });
  }

  if (touches((file) => file.includes("corpus"))) {
    guards.push({
      id: "check:corpus-baselines",
      command: run("check:corpus-baselines"),
      fixHint: "Keep the corpus baselines and lists consistent with the lock.",
    });
  }

  return guards.filter(
    (guard) => guard.command[1] !== "run" || hasScript(scripts, guard.id),
  );
};

const git = (args: readonly string[]): string =>
  execFileSync("git", args, { cwd: REPO_ROOT, encoding: "utf8" });

const lines = (output: string): string[] =>
  output.split("\n").filter((line) => line !== "");

/** The files the branch and its uncommitted changes touch, deleted ones included. */
const changedFiles = (base: string): readonly string[] => {
  const mergeBase = git(["merge-base", base, "HEAD"]).trim();
  return [
    ...new Set([
      ...lines(git(["diff", "--name-only", `${mergeBase}...HEAD`])),
      ...lines(git(["diff", "--name-only", "HEAD"])),
    ]),
  ].toSorted();
};

type Outcome = { guard: PrGuard; ok: boolean; output: string; ms: number };

const runGuard = async (guard: PrGuard): Promise<Outcome> => {
  const started = performance.now();
  const child = Bun.spawn([...guard.command], {
    cwd: REPO_ROOT,
    stdout: "pipe",
    stderr: "pipe",
  });
  const [stdout, stderr, code] = await Promise.all([
    new Response(child.stdout).text(),
    new Response(child.stderr).text(),
    child.exited,
  ]);
  return {
    guard,
    ok: code === 0,
    output: `${stdout}${stderr}`.trim(),
    ms: performance.now() - started,
  };
};

const runAll = async (guards: readonly PrGuard[]): Promise<Outcome[]> => {
  const outcomes: Outcome[] = [];
  let next = 0;
  const worker = async (): Promise<void> => {
    while (next < guards.length) {
      const guard = guards[next];
      next += 1;
      if (guard === undefined) return;
      const outcome = await runGuard(guard);
      const seconds = (outcome.ms / 1000).toFixed(1);
      console.log(`${outcome.ok ? "ok  " : "FAIL"} ${guard.id} (${seconds} s)`);
      outcomes.push(outcome);
    }
  };
  await Promise.all(
    Array.from({ length: Math.min(CONCURRENCY, guards.length) }, worker),
  );
  return outcomes;
};

const main = async (): Promise<number> => {
  const argv = process.argv.slice(2);
  const baseIndex = argv.indexOf("--base");
  const base = baseIndex === -1 ? "origin/main" : argv[baseIndex + 1];
  if (base === undefined) {
    console.error("--base needs a ref");
    return 2;
  }
  const scripts = JSON.parse(
    readFileSync(path.join(REPO_ROOT, "package.json"), "utf8"),
  ).scripts as Record<string, string>;
  const files = changedFiles(base);
  const guards = selectPrGuards(files, scripts);
  if (guards.length === 0) {
    console.log(`No source guard reads the ${files.length} changed file(s).`);
    return 0;
  }
  if (argv.includes("--list")) {
    for (const guard of guards) console.log(guard.id);
    return 0;
  }
  const failed = (await runAll(guards)).filter((outcome) => !outcome.ok);
  for (const { guard, output } of failed) {
    console.error(`\n${guard.id} failed:\n${output}\nFix: ${guard.fixHint}`);
  }
  return failed.length === 0 ? 0 : 1;
};

if (import.meta.main) process.exit(await main());
