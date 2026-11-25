#!/usr/bin/env bun

/**
 * Shrink-only typecheck of test files.
 *
 * Every package's `tsconfig.build.json` leaves out `*.test.ts` and
 * `__tests__`, so `bun run typecheck` never reads test code and type errors
 * pile up there unseen. Each package's `tsconfig.test.json` (or the config
 * named below) covers its tests and their support files, and this guard holds
 * the errors they report against a committed baseline that may only go down:
 * a new error fails, and a fixed one must leave the baseline.
 *
 * An error is identified by file, code and message, never by line, so moving
 * code around does not churn the baseline. Repeats of the same error in one
 * file are counted.
 *
 * Modes:
 *   bun scripts/test-typecheck-baseline.ts                  report per file
 *   bun scripts/test-typecheck-baseline.ts --write-baseline regenerate baseline
 *   bun scripts/test-typecheck-baseline.ts --check          CI gate
 *   ... --project <id>                                      limit to projects
 */

import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const BASELINE_RELATIVE_PATH = "scripts/test-typecheck-baseline.json";
const BASELINE_PATH = path.join(REPO_ROOT, BASELINE_RELATIVE_PATH);
const REGENERATED_PATH = path.join(
  REPO_ROOT,
  ".cache/test-typecheck/test-typecheck-baseline.json",
);
const WRITE_COMMAND = "bun run typecheck:tests:write";

type Compiler = "native" | "vue";

type TestProject = {
  id: string;
  config: string;
  compiler: Compiler;
};

/** One project per package that has tests; `vue` needs vue-tsc for SFCs. */
export const TEST_PROJECTS: readonly TestProject[] = [
  {
    id: "agents",
    config: "packages/agents/tsconfig.test.json",
    compiler: "native",
  },
  { id: "cli", config: "packages/cli/tsconfig.test.json", compiler: "native" },
  {
    id: "core",
    config: "packages/core/tsconfig.test.json",
    compiler: "native",
  },
  {
    id: "docx-core",
    config: "packages/docx-core/tsconfig.test.json",
    compiler: "native",
  },
  {
    id: "editor-web",
    config: "packages/editor-web/tsconfig.test.json",
    compiler: "native",
  },
  // The fonts package config already includes its tests and Bun's types.
  { id: "fonts", config: "packages/fonts/tsconfig.json", compiler: "native" },
  {
    id: "react",
    config: "packages/react/tsconfig.test.json",
    compiler: "native",
  },
  {
    id: "utf8-boundaries",
    config: "packages/utf8-boundaries/tsconfig.test.json",
    compiler: "native",
  },
  { id: "vue", config: "packages/vue/tsconfig.test.json", compiler: "vue" },
  {
    id: "playground-vue",
    config: "packages/playground-vue/test/tsconfig.json",
    compiler: "vue",
  },
];

/** file -> "TSnnnn: message" -> count */
export type ErrorCounts = Record<string, Record<string, number>>;

type Diagnostic = { file: string; line: number; fingerprint: string };

const DIAGNOSTIC_LINE = /^(.+?)\((\d+),\d+\): error (TS\d+): (.*)$/u;

const toRepoPath = (file: string, cwd: string, root: string): string =>
  path.relative(root, path.resolve(cwd, file)).replaceAll("\\", "/");

/**
 * Parse `--pretty false` output: one `file(line,col): error TSnnnn: message`
 * line per diagnostic. The indented lines a chained message adds below it
 * spell out types from elsewhere and change with them, so the fingerprint
 * keeps the headline only. Absolute checkout paths are made repo-relative so
 * the fingerprint is the same on every machine.
 */
export const parseDiagnostics = (
  output: string,
  { cwd = REPO_ROOT, root = REPO_ROOT } = {},
): Diagnostic[] => {
  const rootPrefix = `${root.replaceAll("\\", "/")}/`;
  return output.split("\n").flatMap((raw) => {
    const match = DIAGNOSTIC_LINE.exec(raw.replace(/\r$/u, ""));
    if (match === null) return [];
    const [, file = "", lineNumber = "0", code = "", message = ""] = match;
    return [
      {
        file: toRepoPath(file, cwd, root),
        line: Number(lineNumber),
        fingerprint: `${code}: ${message.replaceAll(rootPrefix, "")}`,
      },
    ];
  });
};

export const countDiagnostics = (
  diagnostics: readonly Diagnostic[],
): ErrorCounts => {
  const counts = new Map<string, Map<string, number>>();
  for (const { file, fingerprint } of diagnostics) {
    const perFile = counts.get(file) ?? new Map<string, number>();
    perFile.set(fingerprint, (perFile.get(fingerprint) ?? 0) + 1);
    counts.set(file, perFile);
  }
  return Object.fromEntries(
    [...counts.entries()]
      .toSorted(([a], [b]) => a.localeCompare(b))
      .map(([file, perFile]) => [
        file,
        Object.fromEntries(
          [...perFile.entries()].toSorted(([a], [b]) => a.localeCompare(b)),
        ),
      ]),
  );
};

export type Comparison = {
  added: { file: string; fingerprint: string; from: number; to: number }[];
  fixed: { file: string; fingerprint: string; from: number; to: number }[];
};

export const compareCounts = (
  baseline: ErrorCounts,
  current: ErrorCounts,
): Comparison => {
  const added: Comparison["added"] = [];
  const fixed: Comparison["fixed"] = [];
  const files = new Set([...Object.keys(baseline), ...Object.keys(current)]);
  for (const file of [...files].toSorted()) {
    const before = baseline[file] ?? {};
    const after = current[file] ?? {};
    const fingerprints = new Set([
      ...Object.keys(before),
      ...Object.keys(after),
    ]);
    for (const fingerprint of [...fingerprints].toSorted()) {
      const from = before[fingerprint] ?? 0;
      const to = after[fingerprint] ?? 0;
      if (to > from) added.push({ file, fingerprint, from, to });
      if (to < from) fixed.push({ file, fingerprint, from, to });
    }
  }
  return { added, fixed };
};

const projectPrefix = ({ config }: TestProject): string =>
  `${path.posix.dirname(config)}/`;

/** Keep only the files a project owns, for a `--project` subset. */
const ownedBy = (counts: ErrorCounts, projects: readonly TestProject[]) => {
  const prefixes = projects.map(projectPrefix);
  return Object.fromEntries(
    Object.entries(counts).filter(([file]) =>
      prefixes.some((prefix) => file.startsWith(prefix)),
    ),
  );
};

const compilerCommand = ({ config, compiler }: TestProject): string[] => {
  if (compiler === "vue") {
    // vue-tsc's Node launcher patches TypeScript for SFCs; run it directly.
    const packageDirectory = config.split("/").slice(0, 2).join("/");
    return [
      path.join(
        REPO_ROOT,
        packageDirectory,
        "node_modules/vue-tsc/bin/vue-tsc.js",
      ),
      "--noEmit",
      "--pretty",
      "false",
      "-p",
      config,
    ];
  }
  return [
    process.execPath,
    "check",
    "--no-pretty",
    "--all",
    `--project=${config}`,
  ];
};

const checkProject = (project: TestProject): Diagnostic[] => {
  const command = compilerCommand(project);
  const spawned = Bun.spawnSync(command, {
    cwd: REPO_ROOT,
    stdout: "pipe",
    stderr: "pipe",
  });
  const output = `${spawned.stdout.toString()}\n${spawned.stderr.toString()}`;
  const diagnostics = parseDiagnostics(output);
  if (spawned.exitCode !== 0 && diagnostics.length === 0) {
    throw new Error(
      `${project.id}: typecheck exited ${spawned.exitCode} without diagnostics.\n${output}`,
    );
  }
  // A test imports other packages' sources through workspace links; their
  // errors belong to the project that owns them, so each is counted once.
  const owned = projectPrefix(project);
  return diagnostics.filter(
    ({ file }) => file.startsWith(owned) && !file.includes("node_modules/"),
  );
};

const total = (counts: ErrorCounts): number =>
  Object.values(counts).reduce(
    (sum, perFile) =>
      sum + Object.values(perFile).reduce((inner, count) => inner + count, 0),
    0,
  );

const readBaseline = (): ErrorCounts => {
  if (!existsSync(BASELINE_PATH)) return {};
  const parsed: unknown = JSON.parse(readFileSync(BASELINE_PATH, "utf8"));
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
    throw new TypeError(`${BASELINE_RELATIVE_PATH} must be an object`);
  }
  return parsed as ErrorCounts;
};

const selectedProjects = (argv: readonly string[]): readonly TestProject[] => {
  const ids = argv.flatMap((argument, index) =>
    argument === "--project" ? [argv[index + 1] ?? ""] : [],
  );
  if (ids.length === 0) return TEST_PROJECTS;
  const unknown = ids.filter(
    (id) => !TEST_PROJECTS.some((project) => project.id === id),
  );
  if (unknown.length > 0) {
    throw new Error(`Unknown test project(s): ${unknown.join(", ")}`);
  }
  return TEST_PROJECTS.filter(({ id }) => ids.includes(id));
};

const main = () => {
  const argv = process.argv.slice(2);
  const mode = argv.includes("--write-baseline")
    ? "write"
    : argv.includes("--check")
      ? "check"
      : "report";
  const projects = selectedProjects(argv);
  const diagnostics = projects.flatMap((project) => {
    console.log(`test-typecheck: checking ${project.id}`);
    return checkProject(project);
  });
  const current = countDiagnostics(diagnostics);
  const subset = projects.length !== TEST_PROJECTS.length;

  if (mode === "write") {
    // A subset rewrites only the files its projects own.
    const kept = subset
      ? Object.fromEntries(
          Object.entries(readBaseline()).filter(
            ([file]) => !(file in ownedBy({ [file]: {} }, projects)),
          ),
        )
      : {};
    const merged = { ...kept, ...current };
    const sorted = Object.fromEntries(
      Object.entries(merged).toSorted(([a], [b]) => a.localeCompare(b)),
    );
    writeFileSync(BASELINE_PATH, `${JSON.stringify(sorted, null, 2)}\n`);
    console.log(
      `Wrote ${total(sorted)} test-file type errors across ${Object.keys(sorted).length} files to ${BASELINE_RELATIVE_PATH}`,
    );
    return 0;
  }

  if (mode === "report") {
    for (const { file, line, fingerprint } of diagnostics) {
      console.log(`${file}:${line} ${fingerprint.split("\n")[0]}`);
    }
    console.log(
      `\n${total(current)} test-file type errors across ${Object.keys(current).length} files`,
    );
    return 0;
  }

  const baseline = subset ? ownedBy(readBaseline(), projects) : readBaseline();
  const { added, fixed } = compareCounts(baseline, current);
  if (!subset && (added.length > 0 || fixed.length > 0)) {
    // CI uploads this so a fixed error can be locked in without a local run.
    mkdirSync(path.dirname(REGENERATED_PATH), { recursive: true });
    writeFileSync(REGENERATED_PATH, `${JSON.stringify(current, null, 2)}\n`);
  }
  if (added.length > 0) {
    console.error("New type errors in test files:");
    for (const { file, fingerprint, from, to } of added) {
      const lines = diagnostics
        .filter(
          (entry) => entry.file === file && entry.fingerprint === fingerprint,
        )
        .map((entry) => entry.line);
      console.error(
        `  ${file}:${lines.join(",")} (${from} -> ${to}) ${fingerprint}`,
      );
    }
    console.error(
      "\nFix the new errors. The baseline only records errors that predate this guard.",
    );
  }
  if (fixed.length > 0) {
    console.error(
      "Type errors fixed in test files (remove them from the baseline):",
    );
    for (const { file, fingerprint, from, to } of fixed) {
      console.error(
        `  ${file} (${from} -> ${to}) ${fingerprint.split("\n")[0]}`,
      );
    }
    console.error(
      `\nLock the win in with \`${WRITE_COMMAND}\`, or copy the ` +
        "`test-typecheck-baseline` CI artifact over the baseline.",
    );
  }
  if (added.length > 0 || fixed.length > 0) return 1;
  console.log(
    `Test typecheck: ${total(current)} baselined errors across ${Object.keys(current).length} files, none added.`,
  );
  return 0;
};

if (import.meta.main) {
  process.exitCode = main();
}
