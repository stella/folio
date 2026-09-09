#!/usr/bin/env bun

import { existsSync, readFileSync, realpathSync, statSync } from "node:fs";
import path from "node:path";

import { TaggedError } from "better-result";
import ts from "typescript";

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const REPORT_VERSION = 1;
const SOURCE_EXTENSIONS = [".ts", ".tsx", ".mts", ".cts", ".js", ".jsx", ".mjs", ".cjs"];
const TEST_FILE_PATTERN = /(?:^|\/)[^/]+(?:\.[^/]+)*\.(?:test|spec)\.[cm]?[jt]sx?$/u;
const TEST_IMPORT_DISTANCE = 2;

const SEAM = {
  kernel: "docx-kernel",
  docxModel: "docx-model",
  parserSerializer: "parser-serializer",
  editableModel: "editable-model",
  flowConversion: "flow-conversion",
  measurement: "measurement-shaping",
  pagination: "pagination",
  displayList: "display-list",
  domRender: "dom-render",
  controller: "controller",
  adapter: "framework-adapter",
  parity: "interoperability-parity",
  tooling: "repository-tooling",
  tests: "test-harness",
  other: "other",
} as const;

type Seam = (typeof SEAM)[keyof typeof SEAM];

type PackageInfo = {
  readonly directory: string;
  readonly name: string;
  readonly published: boolean;
};

export type FileOrientation = {
  readonly file: string;
  readonly package: PackageInfo | null;
  readonly seam: Seam;
  readonly instructions: readonly string[];
  readonly tests: readonly string[];
  readonly imports: readonly string[];
  readonly importedBy: readonly string[];
  readonly checks: readonly string[];
  readonly requiresChangeset: boolean;
};

export type OrientationReport = {
  readonly version: typeof REPORT_VERSION;
  readonly files: readonly FileOrientation[];
  readonly checks: readonly string[];
  readonly changesetPackages: readonly string[];
};

class OrientError extends TaggedError("OrientError")<{ message: string; cause?: unknown }> {}

type CliFlags = {
  readonly json: boolean;
  readonly help: boolean;
  readonly diffRef: string | null;
  readonly paths: readonly string[];
};

const parseArgs = (argv: readonly string[]): CliFlags => {
  let json = false;
  let help = false;
  let diffRef: string | null = null;
  const paths: string[] = [];

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === undefined) continue;
    if (arg === "--json") {
      json = true;
      continue;
    }
    if (arg === "--help" || arg === "-h") {
      help = true;
      continue;
    }
    if (arg === "--diff") {
      const candidate = argv.at(index + 1);
      if (candidate !== undefined && !candidate.startsWith("--")) {
        diffRef = candidate;
        index += 1;
      } else {
        diffRef = "HEAD";
      }
      continue;
    }
    paths.push(arg);
  }

  if (diffRef !== null && paths.length > 0) {
    throw new OrientError({ message: "Use either --diff [ref] or explicit paths, not both." });
  }
  return { json, help, diffRef: diffRef ?? (paths.length === 0 ? "HEAD" : null), paths };
};

const runGit = (repoRoot: string, args: readonly string[]): string => {
  const result = Bun.spawnSync(["git", ...args], {
    cwd: repoRoot,
    stdout: "pipe",
    stderr: "pipe",
  });
  if (result.exitCode !== 0) {
    const detail = result.stderr.toString().trim();
    throw new OrientError({
      message: `git ${args.join(" ")} failed${detail.length === 0 ? "" : `: ${detail}`}`,
    });
  }
  return result.stdout.toString();
};

const gitFiles = (repoRoot: string): readonly string[] =>
  runGit(repoRoot, ["ls-files"])
    .split("\n")
    .filter((file) => file.length > 0)
    .toSorted((a, b) => a.localeCompare(b));

const relativeInsideRepo = (repoRoot: string, candidate: string): string => {
  const absolute = path.resolve(repoRoot, candidate);
  const relative = path.relative(repoRoot, absolute);
  if (relative === "" || relative === ".") return "";
  if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw new OrientError({ message: `Path is outside the repository: ${candidate}` });
  }
  return relative.split(path.sep).join("/");
};

const canonicalPathInsideRepo = (canonicalRepoRoot: string, candidate: string): string => {
  const canonical = realpathSync(candidate);
  const relative = path.relative(canonicalRepoRoot, canonical);
  if (relative === ".." || relative.startsWith(`..${path.sep}`) || path.isAbsolute(relative)) {
    throw new OrientError({ message: `Path resolves outside the repository: ${candidate}` });
  }
  return canonical;
};

type ResolveFilesOptions = {
  readonly repoRoot: string;
  readonly trackedFiles: readonly string[];
  readonly paths: readonly string[];
  readonly diffRef: string | null;
};

export const resolveFiles = ({
  repoRoot,
  trackedFiles,
  paths,
  diffRef,
}: ResolveFilesOptions): readonly string[] => {
  if (diffRef !== null) {
    const tracked = new Set(trackedFiles);
    return runGit(repoRoot, ["diff", "--name-only", "--diff-filter=ACMRTUXB", diffRef])
      .split("\n")
      .filter((file) => tracked.has(file))
      .toSorted((a, b) => a.localeCompare(b));
  }

  const requested = paths.length === 0 ? ["."] : paths;
  const resolved = new Set<string>();
  for (const requestedPath of requested) {
    const relative = relativeInsideRepo(repoRoot, requestedPath);
    const absolute = path.join(repoRoot, relative);
    if (!existsSync(absolute)) {
      throw new OrientError({ message: `Path does not exist: ${requestedPath}` });
    }
    if (!statSync(absolute).isDirectory()) {
      resolved.add(relative);
      continue;
    }
    const prefix = relative.length === 0 ? "" : `${relative}/`;
    for (const file of trackedFiles) {
      if (file.startsWith(prefix)) resolved.add(file);
    }
  }
  return [...resolved].toSorted((a, b) => a.localeCompare(b));
};

export const classifySeam = (file: string): Seam => {
  if (file.startsWith("crates/docx-kernel/") || file.startsWith("packages/docx-core/")) {
    return SEAM.kernel;
  }
  if (file.startsWith("parity/")) return SEAM.parity;
  if (file.startsWith("scripts/")) return SEAM.tooling;
  if (file.startsWith("tests/") || TEST_FILE_PATTERN.test(file)) return SEAM.tests;
  if (/^packages\/(?:react|vue|nuxt|playground|playground-vue)\//u.test(file)) {
    return SEAM.adapter;
  }
  if (file.includes("/layout-bridge/convert/")) return SEAM.flowConversion;
  if (file.includes("/layout-engine/measure/") || file.includes("/shaping/")) {
    return SEAM.measurement;
  }
  if (file.includes("/layout-engine/")) return SEAM.pagination;
  if (file.includes("/display-list/build/")) return SEAM.displayList;
  if (file.includes("/display-list/dom/") || file.includes("/layout-painter/")) {
    return SEAM.domRender;
  }
  if (file.includes("/controller/") || file.includes("/managers/")) return SEAM.controller;
  if (file.includes("/prosemirror/")) return SEAM.editableModel;
  if (file.startsWith("packages/core/src/docx/")) return SEAM.parserSerializer;
  if (file.startsWith("packages/core/src/types/")) return SEAM.docxModel;
  return SEAM.other;
};

const packageFor = (repoRoot: string, file: string): PackageInfo | null => {
  const match = /^(packages\/[^/]+)\//u.exec(file);
  const directory = match?.at(1);
  if (directory === undefined) return null;
  const manifestPath = path.join(repoRoot, directory, "package.json");
  if (!existsSync(manifestPath)) return null;
  const canonicalManifestPath = canonicalPathInsideRepo(realpathSync(repoRoot), manifestPath);
  const manifest: unknown = JSON.parse(readFileSync(canonicalManifestPath, "utf8"));
  if (typeof manifest !== "object" || manifest === null || !("name" in manifest)) return null;
  const name = Reflect.get(manifest, "name");
  const isPrivate = Reflect.get(manifest, "private");
  return typeof name === "string" ? { directory, name, published: isPrivate !== true } : null;
};

const instructionsFor = (repoRoot: string, file: string): readonly string[] => {
  const directories: string[] = [];
  let current = path.dirname(path.join(repoRoot, file));
  while (current === repoRoot || current.startsWith(`${repoRoot}${path.sep}`)) {
    const candidate = path.join(current, "AGENTS.md");
    if (existsSync(candidate)) directories.push(path.relative(repoRoot, candidate));
    if (current === repoRoot) break;
    current = path.dirname(current);
  }
  return directories.reverse().map((filePath) => filePath.split(path.sep).join("/"));
};

const resolveImport = (
  importer: string,
  specifier: string,
  tracked: ReadonlySet<string>,
): string | null => {
  if (!specifier.startsWith(".")) return null;
  const base = path.posix.normalize(path.posix.join(path.posix.dirname(importer), specifier));
  const extension = path.posix.extname(base);
  const candidates =
    extension.length > 0
      ? [
          base,
          ...SOURCE_EXTENSIONS.map(
            (sourceExtension) => base.slice(0, -extension.length) + sourceExtension,
          ),
        ]
      : [
          ...SOURCE_EXTENSIONS.map((sourceExtension) => `${base}${sourceExtension}`),
          ...SOURCE_EXTENSIONS.map((sourceExtension) => `${base}/index${sourceExtension}`),
        ];
  return candidates.find((candidate) => tracked.has(candidate)) ?? null;
};

type ImportGraph = {
  readonly imports: ReadonlyMap<string, readonly string[]>;
  readonly importedBy: ReadonlyMap<string, readonly string[]>;
};

export const buildImportGraph = (
  repoRoot: string,
  trackedFiles: readonly string[],
): ImportGraph => {
  const sourceFiles = trackedFiles.filter((file) => SOURCE_EXTENSIONS.includes(path.extname(file)));
  const tracked = new Set(trackedFiles);
  const canonicalRepoRoot = realpathSync(repoRoot);
  const imports = new Map<string, readonly string[]>();
  const reverse = new Map<string, string[]>();

  for (const file of sourceFiles) {
    const sourcePath = canonicalPathInsideRepo(canonicalRepoRoot, path.join(repoRoot, file));
    const source = readFileSync(sourcePath, "utf8");
    const resolved = ts
      .preProcessFile(source, true, true)
      .importedFiles.map(({ fileName }) => resolveImport(file, fileName, tracked))
      .filter((candidate): candidate is string => candidate !== null);
    const unique = [...new Set(resolved)].toSorted((a, b) => a.localeCompare(b));
    imports.set(file, unique);
    for (const dependency of unique) {
      const importers = reverse.get(dependency) ?? [];
      importers.push(file);
      reverse.set(dependency, importers);
    }
  }

  return {
    imports,
    importedBy: new Map(
      [...reverse].map(([file, importers]) => [
        file,
        [...new Set(importers)].toSorted((a, b) => a.localeCompare(b)),
      ]),
    ),
  };
};

const testsFor = (
  file: string,
  trackedFiles: readonly string[],
  importedBy: ReadonlyMap<string, readonly string[]>,
): readonly string[] => {
  if (TEST_FILE_PATTERN.test(file)) return [file];
  const directory = path.posix.dirname(file);
  const stem = path.posix.basename(file).replace(/\.[^.]+$/u, "");
  const adjacentPrefix = `${directory}/${stem}`;
  const adjacent = trackedFiles.filter(
    (candidate) => TEST_FILE_PATTERN.test(candidate) && candidate.startsWith(adjacentPrefix),
  );
  const related = new Set<string>();
  const visited = new Set([file]);
  let frontier = [file];
  for (let distance = 0; distance < TEST_IMPORT_DISTANCE; distance += 1) {
    const next: string[] = [];
    for (const dependency of frontier) {
      for (const importer of importedBy.get(dependency) ?? []) {
        if (visited.has(importer)) continue;
        visited.add(importer);
        if (TEST_FILE_PATTERN.test(importer)) {
          related.add(importer);
          continue;
        }
        next.push(importer);
      }
    }
    frontier = next;
  }
  return [...new Set([...adjacent, ...related])].toSorted((a, b) => a.localeCompare(b));
};

export const checksForFile = (file: string): readonly string[] => {
  const checks = new Set<string>();
  if (file.startsWith("packages/core/")) {
    checks.add("bun --filter @stll/folio-core test");
    checks.add("bun --filter @stll/folio-core typecheck");
  }
  if (
    file.startsWith("packages/core/src/layout-") ||
    file.startsWith("packages/core/src/display-list/") ||
    file.startsWith("packages/core/src/paged-layout/")
  ) {
    checks.add("bun run test:interactions");
    checks.add("bun run test:measure-parity");
  }
  if (file.startsWith("packages/docx-core/") || file.startsWith("crates/docx-kernel/")) {
    checks.add("bun run check:docx-kernel");
  }
  if (/^packages\/(?:react|vue)\/src\//u.test(file)) {
    checks.add("bun run check:parity-contract");
    checks.add("bun run check:export-parity");
  }
  if (file.startsWith("parity/")) {
    checks.add("bun run typecheck:parity");
    checks.add("bun test parity/__tests__");
  }
  if (file.startsWith("scripts/")) {
    checks.add("bun test scripts");
    checks.add("bun run typecheck:tooling");
  }
  checks.add("bun run lint");
  checks.add("bun run format:check");
  return [...checks];
};

export const buildReport = (
  repoRoot: string,
  trackedFiles: readonly string[],
  files: readonly string[],
): OrientationReport => {
  const graph = buildImportGraph(repoRoot, trackedFiles);
  const orientations = files.map((file): FileOrientation => {
    const packageInfo = packageFor(repoRoot, file);
    const requiresChangeset =
      packageInfo?.published === true && file.startsWith(`${packageInfo.directory}/src/`);
    return {
      file,
      package: packageInfo,
      seam: classifySeam(file),
      instructions: instructionsFor(repoRoot, file),
      tests: testsFor(file, trackedFiles, graph.importedBy),
      imports: graph.imports.get(file) ?? [],
      importedBy: graph.importedBy.get(file) ?? [],
      checks: checksForFile(file),
      requiresChangeset,
    };
  });
  return {
    version: REPORT_VERSION,
    files: orientations,
    checks: [...new Set(orientations.flatMap(({ checks }) => checks))],
    changesetPackages: [
      ...new Set(
        orientations.flatMap(({ package: packageInfo, requiresChangeset }) =>
          requiresChangeset && packageInfo !== null ? [packageInfo.name] : [],
        ),
      ),
    ].toSorted((a, b) => a.localeCompare(b)),
  };
};

const section = (label: string, values: readonly string[]): string =>
  values.length === 0
    ? `${label}: none`
    : `${label}:\n${values.map((value) => `  - ${value}`).join("\n")}`;

const printHuman = (report: OrientationReport): void => {
  for (const file of report.files) {
    console.log(`\n${file.file}`);
    console.log(`  seam: ${file.seam}`);
    console.log(`  package: ${file.package?.name ?? "repository"}`);
    console.log(`  changeset: ${file.requiresChangeset ? "required" : "not required"}`);
    console.log(section("  instructions", file.instructions));
    console.log(section("  focused tests", file.tests));
    console.log(section("  imports", file.imports));
    console.log(section("  imported by", file.importedBy));
  }
  console.log(`\n${section("Checks", report.checks)}`);
  console.log(section("Changeset packages", report.changesetPackages));
};

const HELP = `Repository orientation

Usage:
  bun run orient -- path/to/file.ts [more paths]
  bun run orient -- --diff [ref]
  bun run orient -- --diff main --json

Reports architectural ownership, nearby tests, direct source relationships,
required checks, and release impact. With no arguments, reports the current
tracked diff against HEAD.
`;

const main = (): void => {
  const flags = parseArgs(Bun.argv.slice(2));
  if (flags.help) {
    console.log(HELP);
    return;
  }
  const trackedFiles = gitFiles(REPO_ROOT);
  const files = resolveFiles({
    repoRoot: REPO_ROOT,
    trackedFiles,
    paths: flags.paths,
    diffRef: flags.diffRef,
  });
  if (files.length === 0) {
    throw new OrientError({ message: "No tracked files matched the request." });
  }
  const report = buildReport(REPO_ROOT, trackedFiles, files);
  if (flags.json) {
    console.log(JSON.stringify(report, null, 2));
    return;
  }
  printHuman(report);
};

if (import.meta.main) {
  try {
    main();
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    process.stderr.write(`orient: ${message}\n`);
    process.exitCode = 1;
  }
}
