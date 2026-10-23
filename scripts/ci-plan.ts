import { readFileSync } from "node:fs";
import path from "node:path";

import { changedApiPaths } from "./ci-api-plan";

const TYPECHECK_ALL_INPUTS = new Set([
  ".github/workflows/ci.yml",
  "bun.lock",
  "bunfig.toml",
  "package.json",
  "scripts/ci-plan.ts",
  "scripts/ci-plan.test.ts",
  "tsconfig.base.json",
]);

export type TypecheckProject = {
  name: string;
  command: readonly string[];
};

export type TypecheckPlan =
  { type: "all" } | { type: "selected"; projects: readonly TypecheckProject[] };

type Manifest = {
  name?: unknown;
  scripts?: unknown;
  dependencies?: unknown;
  devDependencies?: unknown;
  peerDependencies?: unknown;
};

const readManifest = (file: string): Manifest =>
  JSON.parse(readFileSync(file, "utf8"));

const rootProjects = [
  {
    name: "typecheck:benchmarks",
    tsconfig: "benchmarks/tsconfig.json",
    dependencies: ["@stll/folio-core"],
  },
  {
    name: "typecheck:parity",
    tsconfig: "parity/tsconfig.check.json",
    dependencies: ["@stll/folio-core"],
  },
  {
    name: "typecheck:corpus",
    tsconfig: "scripts/tsconfig.corpus.json",
    dependencies: ["@stll/docx-core", "@stll/folio-core"],
  },
  {
    name: "typecheck:corpus-edit-fuzz",
    tsconfig: "test/tsconfig.corpus-edit-fuzz.json",
    dependencies: ["@stll/folio-core"],
  },
  {
    name: "typecheck:tooling",
    tsconfig: "scripts/tsconfig.budgets.json",
    dependencies: ["@stll/docx-core", "@stll/folio-core"],
  },
  {
    name: "typecheck:reserved-values",
    tsconfig: "specifications/reserved-values/tsconfig.json",
    dependencies: ["@stll/docx-core"],
  },
  {
    name: "typecheck:container-contract",
    tsconfig: "specifications/container-contract/tsconfig.json",
    dependencies: [],
  },
] as const;

type TsconfigInputs = { include?: unknown; files?: unknown; exclude?: unknown };

const rootProjectMatches = (
  tsconfigPath: string,
  file: string,
  repoRoot: string,
) => {
  if (file === tsconfigPath) return true;
  const config: TsconfigInputs = JSON.parse(
    readFileSync(path.join(repoRoot, tsconfigPath), "utf8"),
  );
  const directory = path.posix.dirname(tsconfigPath);
  const inputs = [config.include, config.files].flatMap((value) =>
    Array.isArray(value)
      ? value.filter((entry): entry is string => typeof entry === "string")
      : [],
  );
  const excludes = Array.isArray(config.exclude)
    ? config.exclude.filter(
        (entry): entry is string => typeof entry === "string",
      )
    : [];
  return inputs.some((input) => {
    const relativePattern = input === "." ? "**" : input;
    const pattern = path.posix.normalize(
      path.posix.join(directory, relativePattern),
    );
    if (!new Bun.Glob(pattern).match(file)) return false;
    return !excludes.some((exclude) => {
      const excludedPattern = exclude.endsWith("/")
        ? `${exclude}**`
        : `${exclude}/**`;
      return new Bun.Glob(
        path.posix.normalize(path.posix.join(directory, excludedPattern)),
      ).match(file);
    });
  });
};

const isTypecheckedInput = (file: string) =>
  !file.endsWith(".md") &&
  !/(?:^|\/)(?:__tests__|test|tests)\//u.test(file) &&
  !/\.(?:test|spec)\.[^/]+$/u.test(file);

/** Select the projects used by both the CI Typecheck job and local affected checks. */
export const typecheckPlanForPaths = (
  changedPaths: readonly string[],
  repoRoot = path.resolve(import.meta.dir, ".."),
): TypecheckPlan => {
  if (changedPaths.some((file) => TYPECHECK_ALL_INPUTS.has(file)))
    return { type: "all" };

  const packageDirectories = [
    ...new Bun.Glob("packages/*/package.json").scanSync({
      cwd: repoRoot,
      onlyFiles: true,
    }),
  ].toSorted();
  const packages = packageDirectories.flatMap((manifestPath) => {
    const manifest = readManifest(path.join(repoRoot, manifestPath));
    if (
      typeof manifest.name !== "string" ||
      typeof manifest.scripts !== "object" ||
      manifest.scripts === null ||
      !("typecheck" in manifest.scripts)
    )
      return [];
    const dependencies = [
      manifest.dependencies,
      manifest.devDependencies,
      manifest.peerDependencies,
    ].flatMap((value) =>
      typeof value === "object" && value !== null ? Object.keys(value) : [],
    );
    return [
      {
        name: manifest.name,
        directory: path.posix.dirname(manifestPath),
        dependencies,
      },
    ];
  });
  const selectedNames = new Set(
    packages
      .filter(({ directory }) =>
        changedPaths.some(
          (file) =>
            isTypecheckedInput(file) &&
            (file === directory || file.startsWith(`${directory}/`)),
        ),
      )
      .map(({ name }) => name),
  );
  let previousSize = -1;
  while (selectedNames.size !== previousSize) {
    previousSize = selectedNames.size;
    for (const project of packages) {
      if (
        project.dependencies.some((dependency) => selectedNames.has(dependency))
      )
        selectedNames.add(project.name);
    }
  }

  const projects: TypecheckProject[] = packages
    .filter(({ name }) => selectedNames.has(name))
    .map(({ name }) => ({
      name,
      command: ["bun", "--filter", name, "typecheck"],
    }));
  for (const project of rootProjects) {
    if (
      project.dependencies.some((dependency) =>
        selectedNames.has(dependency),
      ) ||
      changedPaths.some((file) =>
        rootProjectMatches(project.tsconfig, file, repoRoot),
      )
    )
      projects.push({
        name: project.name,
        command: ["bun", "run", project.name],
      });
  }
  return { type: "selected", projects };
};

type ChangedTypecheckPathsOptions = {
  repoRoot: string;
  base: string;
  head: string;
  fallbackOnMissingBase?: boolean;
};

export const changedTypecheckPaths = ({
  repoRoot,
  base,
  head,
  fallbackOnMissingBase = false,
}: ChangedTypecheckPathsOptions): string[] => {
  if (fallbackOnMissingBase) {
    const resolved = Bun.spawnSync(
      ["git", "rev-parse", "--verify", `${base}^{commit}`],
      {
        cwd: repoRoot,
        stdout: "pipe",
        stderr: "pipe",
      },
    );
    if (resolved.exitCode !== 0) return ["package.json"];
  }
  return changedApiPaths({ repoRoot, base, head });
};

type RunTypecheckPlanOptions = {
  run: (command: readonly string[]) => number;
};

export const runTypecheckPlan = (
  plan: TypecheckPlan,
  { run }: RunTypecheckPlanOptions,
): number => {
  if (plan.type === "all") {
    process.stdout.write("Typechecking all projects.\n");
    return run(["bun", "run", "typecheck"]);
  }
  if (plan.projects.length === 0) {
    process.stdout.write("Skipped typecheck: no typechecked inputs changed.\n");
    return 0;
  }
  process.stdout.write(
    `Typechecking: ${plan.projects.map(({ name }) => name).join(", ")}\n`,
  );
  for (const { command } of plan.projects) {
    const exitCode = run(command);
    if (exitCode !== 0) return exitCode;
  }
  return 0;
};

const run = (command: readonly string[]) => {
  process.stdout.write(`Running: ${command.join(" ")}\n`);
  return Bun.spawnSync(command, {
    cwd: process.cwd(),
    stdin: "inherit",
    stdout: "inherit",
    stderr: "inherit",
  }).exitCode;
};

if (import.meta.main) {
  const args = process.argv.slice(2);
  if (args.length !== 4 || args.at(0) !== "--base" || args.at(2) !== "--head")
    throw new TypeError(
      "Usage: bun scripts/ci-plan.ts --base <sha> --head <sha>",
    );
  const repoRoot = path.resolve(import.meta.dir, "..");
  const base = args[1];
  const head = args[3];
  if (!base || !head)
    throw new TypeError("Typecheck planning requires a base and head.");
  const changedPaths = /^0+$/u.test(base)
    ? ["package.json"]
    : changedTypecheckPaths({
        repoRoot,
        base,
        head,
        fallbackOnMissingBase: process.env["TYPECHECK_PUSH"] === "true",
      });
  process.exitCode = runTypecheckPlan(
    typecheckPlanForPaths(changedPaths, repoRoot),
    { run },
  );
}
