#!/usr/bin/env bun
/** Derive the PR API gate from workspace manifests, never a package allowlist. */
import { appendFileSync, readFileSync } from "node:fs";
import path from "node:path";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const API_CONTROL_PATHS = new Set([
  "scripts/api-surface-budget.json",
  "scripts/api-surface-budget.ts",
  ".github/workflows/ci.yml",
  "scripts/ci-api-plan.ts",
  "scripts/ci-api-plan.test.ts",
  "scripts/ci-run-contract.json",
]);

const readManifest = (file: string) => {
  const value: unknown = JSON.parse(readFileSync(file, "utf8"));
  if (typeof value !== "object" || value === null || Array.isArray(value))
    throw new TypeError(`Manifest must be an object: ${file}`);
  return value;
};

const workspacePattern = (value: unknown) => {
  if (typeof value !== "string" || value === "" || value.includes("\0") || value.includes("\\"))
    throw new TypeError("Workspace patterns must be nonempty repository-relative strings.");
  const source = value.startsWith("!") ? value.slice(1) : value;
  if (source === "" || path.posix.isAbsolute(source) || source.split("/").includes(".."))
    throw new TypeError(`Invalid workspace pattern: ${value}`);
  const pattern = path.posix.normalize(source);
  if (pattern === ".") throw new TypeError("The repository root cannot be its own workspace.");
  return { pattern: pattern.replace(/\/$/u, ""), excluded: value.startsWith("!") };
};

/** Non-private workspace directories, including newly added publication packages. */
export const publishedPackagePaths = (repoRoot = REPO_ROOT): string[] => {
  const root = readManifest(path.join(repoRoot, "package.json"));
  if (!("workspaces" in root) || !Array.isArray(root.workspaces) || root.workspaces.length === 0)
    throw new TypeError("The root manifest must declare workspace patterns.");
  const patterns = root.workspaces.map(workspacePattern);
  const includes = patterns.filter(({ excluded }) => !excluded);
  const excludes = patterns
    .filter(({ excluded }) => excluded)
    .map(({ pattern }) => new Bun.Glob(pattern));
  if (includes.length === 0)
    throw new TypeError("Workspace patterns must include package directories.");
  const manifests = new Set<string>();
  for (const { pattern } of includes) {
    const matches = [
      ...new Bun.Glob(`${pattern}/package.json`).scanSync({
        cwd: repoRoot,
        onlyFiles: true,
        followSymlinks: false,
      }),
    ];
    if (matches.length === 0) throw new TypeError(`Workspace pattern has no manifests: ${pattern}`);
    for (const manifest of matches) {
      if (!excludes.some((glob) => glob.match(path.posix.dirname(manifest))))
        manifests.add(manifest);
    }
  }
  const names = new Set<string>();
  const published: string[] = [];
  for (const manifestPath of [...manifests].toSorted()) {
    const manifest = readManifest(path.join(repoRoot, manifestPath));
    if (!("name" in manifest) || typeof manifest.name !== "string" || manifest.name.trim() === "")
      throw new TypeError(`Workspace manifest has no package name: ${manifestPath}`);
    if (names.has(manifest.name))
      throw new TypeError(`Duplicate workspace package name: ${manifest.name}`);
    names.add(manifest.name);
    if ("private" in manifest && typeof manifest.private !== "boolean")
      throw new TypeError(`Workspace private field must be boolean: ${manifestPath}`);
    if (!("private" in manifest) || !manifest.private)
      published.push(path.posix.dirname(manifestPath));
  }
  return published;
};

const validateChangedPath = (file: string) => {
  if (
    file === "" ||
    file.includes("\0") ||
    file.includes("\\") ||
    path.posix.isAbsolute(file) ||
    path.posix.normalize(file) !== file ||
    file === "." ||
    file === ".." ||
    file.startsWith("../")
  )
    throw new TypeError(`Invalid repository-relative changed path: ${JSON.stringify(file)}`);
};

type ApiRequiredForPathsOptions = {
  changedPaths: readonly string[];
  publishedPackagePaths: readonly string[];
};
export const apiRequiredForPaths = ({
  changedPaths,
  publishedPackagePaths: packages,
}: ApiRequiredForPathsOptions): boolean => {
  // Validate the whole input before an early match can hide malformed paths.
  for (const file of changedPaths) validateChangedPath(file);
  for (const directory of packages) validateChangedPath(directory);
  return changedPaths.some((file) => {
    // Manifests can change workspace membership or publication status, including removal.
    if (file === "package.json" || file.endsWith("/package.json") || API_CONTROL_PATHS.has(file))
      return true;
    return packages.some(
      (directory) => file === `${directory}/src` || file.startsWith(`${directory}/src/`),
    );
  });
};

type DiffOptions = { repoRoot: string; base: string; head: string };
export const changedApiPaths = ({ repoRoot, base, head }: DiffOptions): string[] => {
  if (![base, head].every((sha) => /^(?:[\da-f]{40}|[\da-f]{64})$/iu.test(sha)))
    throw new TypeError("API planning requires complete base and head commit SHAs.");
  const diff = Bun.spawnSync(
    [
      "git",
      "--no-pager",
      "diff",
      "--no-ext-diff",
      "--name-only",
      "--no-renames",
      "-z",
      `${base}...${head}`,
      "--",
    ],
    { cwd: repoRoot, stdout: "pipe", stderr: "pipe" },
  );
  if (diff.exitCode !== 0)
    throw new TypeError(`Cannot read PR API diff: ${diff.stderr.toString().trim()}`);
  const names = new TextDecoder("utf-8", { fatal: true }).decode(diff.stdout);
  if (names === "") return [];
  if (!names.endsWith("\0")) throw new TypeError("Git API diff is not NUL-terminated.");
  const files = names.slice(0, -1).split("\0");
  for (const file of files) validateChangedPath(file);
  return files;
};

const parseArguments = (args: readonly string[]) => {
  let base: string | undefined;
  let head: string | undefined;
  for (let index = 0; index < args.length; index += 2) {
    const flag = args.at(index);
    const value = args.at(index + 1);
    if (!value) throw new TypeError(`Missing value for ${flag ?? "argument"}.`);
    switch (flag) {
      case "--base":
        if (base !== undefined) throw new TypeError("Duplicate --base argument.");
        base = value;
        break;
      case "--head":
        if (head !== undefined) throw new TypeError("Duplicate --head argument.");
        head = value;
        break;
      default:
        throw new TypeError(`Unknown API planner argument: ${flag}`);
    }
  }
  if (!base || !head)
    throw new TypeError("Usage: bun scripts/ci-api-plan.ts --base <sha> --head <sha>");
  return { base, head };
};

const emit = (required: boolean) => {
  const output = `api_required=${String(required)}\n`;
  const destination = process.env["GITHUB_OUTPUT"];
  if (destination) appendFileSync(destination, output);
  process.stdout.write(output);
};

if (import.meta.main) {
  // Configuration, Git and transport failures make the plan required and fail the job.
  try {
    const { base, head } = parseArguments(process.argv.slice(2));
    const root = Bun.spawnSync(["git", "rev-parse", "--show-toplevel"], {
      stdout: "pipe",
      stderr: "pipe",
    });
    if (root.exitCode !== 0)
      throw new TypeError(`Cannot locate API planner repository: ${root.stderr.toString().trim()}`);
    const repoRoot = root.stdout.toString().trim();
    emit(
      apiRequiredForPaths({
        changedPaths: changedApiPaths({ repoRoot, base, head }),
        publishedPackagePaths: publishedPackagePaths(repoRoot),
      }),
    );
  } catch (error) {
    process.exitCode = 1;
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    emit(true);
  }
}
