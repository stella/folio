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
  | { type: "all" }
  | { type: "selected"; projects: readonly TypecheckProject[] };

type Manifest = {
  name?: unknown;
  scripts?: unknown;
  dependencies?: unknown;
  devDependencies?: unknown;
  peerDependencies?: unknown;
};

const readManifest = (file: string): Manifest => JSON.parse(readFileSync(file, "utf8"));

const rootProjects = [
  { name: "typecheck:benchmarks", prefix: "benchmarks/" },
  { name: "typecheck:parity", prefix: "parity/" },
  { name: "typecheck:corpus", prefix: "corpus/" },
  { name: "typecheck:corpus-edit-fuzz", prefix: "test/corpus-edit-fuzz" },
  { name: "typecheck:tooling", prefix: "scripts/" },
  { name: "typecheck:reserved-values", prefix: "specifications/reserved-values/" },
  { name: "typecheck:container-contract", prefix: "specifications/container-contract/" },
] as const;

const rootProjectMatches = (name: (typeof rootProjects)[number]["name"], file: string) => {
  if (name === "typecheck:corpus")
    return (
      file === "scripts/tsconfig.corpus.json" ||
      file.startsWith("scripts/corpus-") ||
      file.startsWith("scripts/lib/corpus-") ||
      file.startsWith("scripts/lib/corpus-invariants/") ||
      file.startsWith("corpus/")
    );
  return file.startsWith(rootProjects.find((project) => project.name === name)?.prefix ?? "\0");
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
  if (changedPaths.some((file) => TYPECHECK_ALL_INPUTS.has(file))) return { type: "all" };

  const packageDirectories = [
    ...new Bun.Glob("packages/*/package.json").scanSync({ cwd: repoRoot, onlyFiles: true }),
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
    ].flatMap((value) => (typeof value === "object" && value !== null ? Object.keys(value) : []));
    return [{ name: manifest.name, directory: path.posix.dirname(manifestPath), dependencies }];
  });
  const selectedNames = new Set(
    packages
      .filter(({ directory }) =>
        changedPaths.some(
          (file) =>
            isTypecheckedInput(file) && (file === directory || file.startsWith(`${directory}/`)),
        ),
      )
      .map(({ name }) => name),
  );
  let previousSize = -1;
  while (selectedNames.size !== previousSize) {
    previousSize = selectedNames.size;
    for (const project of packages) {
      if (project.dependencies.some((dependency) => selectedNames.has(dependency)))
        selectedNames.add(project.name);
    }
  }

  const projects: TypecheckProject[] = packages
    .filter(({ name }) => selectedNames.has(name))
    .map(({ name }) => ({ name, command: ["bun", "--filter", name, "typecheck"] }));
  for (const project of rootProjects) {
    if (
      changedPaths.some(
        (file) => isTypecheckedInput(file) && rootProjectMatches(project.name, file),
      )
    )
      projects.push({ name: project.name, command: ["bun", "run", project.name] });
  }
  return { type: "selected", projects };
};

type RunTypecheckPlanOptions = {
  run: (command: readonly string[]) => number;
};

export const runTypecheckPlan = (plan: TypecheckPlan, { run }: RunTypecheckPlanOptions): number => {
  if (plan.type === "all") {
    process.stdout.write("Typechecking all projects.\n");
    return run(["bun", "run", "typecheck"]);
  }
  if (plan.projects.length === 0) {
    process.stdout.write("Skipped typecheck: no typechecked inputs changed.\n");
    return 0;
  }
  process.stdout.write(`Typechecking: ${plan.projects.map(({ name }) => name).join(", ")}\n`);
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
    throw new TypeError("Usage: bun scripts/ci-plan.ts --base <sha> --head <sha>");
  const repoRoot = path.resolve(import.meta.dir, "..");
  const base = args[1];
  const head = args[3];
  if (!base || !head) throw new TypeError("Typecheck planning requires a base and head.");
  const changedPaths = /^0+$/u.test(base)
    ? ["package.json"]
    : changedApiPaths({ repoRoot, base, head });
  process.exitCode = runTypecheckPlan(typecheckPlanForPaths(changedPaths, repoRoot), { run });
}
