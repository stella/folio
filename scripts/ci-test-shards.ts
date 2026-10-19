#!/usr/bin/env bun
// Refresh ci-test-timings.json from per-file group spans in a successful full-suite log.
// Keep millisecond durations above one second; unmeasured files still run.
import { panic } from "better-result";
import { existsSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import focusedTests from "./ci-focused-tests.json";
import timings from "./ci-test-timings.json";

export const TEST_SHARD_COUNT = 4;
const REPO_ROOT = path.resolve(import.meta.dir, "..");
const TEST_FILE = /(?:\.test|\.spec|_test|_spec)\.(?:[cm]?[jt]sx?)$/u;

type TestSuite = { cwd: string; preloads: string[]; files: string[] };

export const discoverTestSuites = (repoRoot = REPO_ROOT): TestSuite[] => {
  const discoverFiles = (cwd: string, roots: string[]) => {
    const files = new Set<string>();
    for (const root of roots) {
      for (const file of new Bun.Glob(`${root}/**/*`).scanSync({
        cwd: path.join(repoRoot, cwd),
        onlyFiles: true,
      })) {
        if (TEST_FILE.test(file)) files.add(path.posix.join(cwd, file));
      }
    }
    return [...files].toSorted();
  };

  const suites: TestSuite[] = [];
  // The root command runs every packages/* workspace with a test script.
  for (const manifest of new Bun.Glob("packages/*/package.json").scanSync({ cwd: repoRoot })) {
    const pkg = JSON.parse(readFileSync(path.join(repoRoot, manifest), "utf8"));
    const command = pkg.scripts?.test;
    if (command === undefined) continue;
    if (typeof command !== "string" || !command.startsWith("bun test ")) {
      panic(`Unsupported test command in ${manifest}: ${command}`);
    }
    const cwd = path.dirname(manifest);
    const args = command.slice("bun test ".length).split(/\s+/u);
    const preloads: string[] = [];
    const roots: string[] = [];
    for (let i = 0; i < args.length; i++) {
      const arg = args.at(i);
      if (arg === "--preload") {
        const preload = args.at(++i);
        if (!preload) panic(`Missing preload in ${manifest}`);
        preloads.push(preload);
        continue;
      }
      if (!arg || arg.startsWith("-") || !/^[\w./-]+$/u.test(arg)) {
        panic(`Unsupported test argument in ${manifest}: ${arg}`);
      }
      roots.push(arg);
    }
    const bunfig = path.join(repoRoot, cwd, "bunfig.toml");
    if (existsSync(bunfig)) {
      const config: unknown = Bun.TOML.parse(readFileSync(bunfig, "utf8"));
      if (typeof config === "object" && config && "test" in config) {
        const test = config.test;
        if (typeof test === "object" && test && "preload" in test) {
          const configured = test.preload;
          if (
            !Array.isArray(configured) ||
            configured.some((preload) => typeof preload !== "string")
          ) {
            panic(`Unsupported test preload in ${path.posix.join(cwd, "bunfig.toml")}`);
          }
          preloads.push(...configured);
        }
      }
    }
    suites.push({ cwd, preloads, files: discoverFiles(cwd, roots) });
  }
  for (const root of ["scripts", "benchmarks/compare"]) {
    suites.push({ cwd: ".", preloads: [], files: discoverFiles(".", [root]) });
  }
  return suites.toSorted((a, b) => a.cwd.localeCompare(b.cwd));
};

// Missing measurements get a small scheduling estimate, never exclusion.
export const assignTestShards = (files: string[]) => {
  const measured = new Map(Object.entries(timings));
  const weight = (file: string) => measured.get(file) ?? 100;
  const loads = Array.from({ length: TEST_SHARD_COUNT }, () => 0);
  const shards = Array.from({ length: TEST_SHARD_COUNT }, () => new Set<string>());
  const ordered = [...new Set(files)].toSorted(
    (a, b) => weight(b) - weight(a) || a.localeCompare(b),
  );
  for (const file of ordered) {
    const shard = loads.indexOf(Math.min(...loads));
    const target = shards.at(shard);
    const load = loads.at(shard);
    if (!target || load === undefined) panic("Test shard unavailable");
    target.add(file);
    loads[shard] = load + weight(file);
  }
  return shards;
};

export const focusedTestFiles = focusedTests;
const SOURCE_FILE = /^packages\/[^/]+\/src\/.*\.[cm]?[jt]sx?$/u;
const CODE_FILE = /\.[cm]?[jt]sx?$/u;
const CODE_EXTENSIONS = ["", ".ts", ".tsx", ".js", ".jsx", ".mts", ".cts", ".mjs", ".cjs"];
const GRAPH_ROOTS = ["packages/*/src", "packages/*/test", "scripts", "benchmarks/compare", "test"];
const FULL_SELECTION_RATIO = 0.7;

type FullTestSuitesOptions = {
  suites: TestSuite[];
  focusedFiles?: string[];
};

/** Full selection includes fixed-list tests outside workspace discovery exactly once. */
export const selectFullTestSuites = ({
  suites,
  focusedFiles = focusedTests,
}: FullTestSuitesOptions): TestSuite[] => {
  const discovered = new Set(suites.flatMap(({ files }) => files));
  const extra = [...new Set(focusedFiles)].filter((file) => !discovered.has(file));
  if (extra.length === 0) return suites;
  return [...suites, { cwd: ".", preloads: [], files: extra }];
};

/** Deleted paths never expand a fast suite; git returns rename destinations. */
export const changedTestPaths = (base: string, repoRoot = REPO_ROOT): string[] => {
  const result = Bun.spawnSync(
    ["git", "diff", "--name-only", "--diff-filter=ACMRT", "-z", `${base}...HEAD`, "--"],
    { cwd: repoRoot, stdout: "pipe", stderr: "pipe" },
  );
  if (result.exitCode !== 0) panic(`Cannot select changed tests: ${result.stderr.toString()}`);
  return result.stdout
    .toString()
    .split("\0")
    .filter((file) => file !== "" && existsSync(path.join(repoRoot, file)));
};

type FastTestSuitesOptions = {
  suites: TestSuite[];
  changedFiles: string[];
  focusedFiles?: string[];
  repoRoot?: string;
  graph?: ModuleGraph;
};

export type ModuleGraph = {
  reverse: Map<string, Set<string>>;
  unresolved: Map<string, string[]>;
  codeFiles: Set<string>;
};

const normalize = (file: string) => file.split(path.sep).join("/");

const resolveFile = (candidate: string) => {
  for (const suffix of CODE_EXTENSIONS) {
    const file = `${candidate}${suffix}`;
    if (existsSync(file) && statSync(file).isFile()) return file;
  }
  for (const extension of CODE_EXTENSIONS.slice(1)) {
    const file = path.join(candidate, `index${extension}`);
    if (existsSync(file) && statSync(file).isFile()) return file;
  }
  if (existsSync(candidate) && statSync(candidate).isFile()) return candidate;
};

type Workspace = { directory: string; exports: Record<string, string | null>; source?: string };

const readWorkspaces = (repoRoot: string) => {
  const workspaces = new Map<string, Workspace>();
  for (const manifest of new Bun.Glob("packages/*/package.json").scanSync({ cwd: repoRoot })) {
    const pkg = JSON.parse(readFileSync(path.join(repoRoot, manifest), "utf8"));
    if (typeof pkg.name !== "string") continue;
    workspaces.set(pkg.name, {
      directory: path.dirname(manifest),
      exports: typeof pkg.exports === "object" && pkg.exports ? pkg.exports : {},
      source: typeof pkg.source === "string" ? pkg.source : undefined,
    });
  }
  return workspaces;
};

const exportedTarget = (workspace: Workspace, subpath: string) => {
  const key = subpath === "" ? "." : `./${subpath}`;
  const exact = workspace.exports[key];
  if (typeof exact === "string") return exact;
  for (const [pattern, target] of Object.entries(workspace.exports)) {
    if (typeof target !== "string" || !pattern.includes("*")) continue;
    const [prefix, suffix = ""] = pattern.split("*");
    if (!key.startsWith(prefix) || !key.endsWith(suffix)) continue;
    const match = key.slice(prefix.length, key.length - suffix.length);
    return target.replace("*", match);
  }
  if (key === ".") return workspace.source;
};

const resolveSpecifier = (
  specifier: string,
  importer: string,
  repoRoot: string,
  workspaces: Map<string, Workspace>,
) => {
  if (specifier.startsWith(".") || specifier.startsWith("/")) {
    const absolute = resolveFile(
      path.resolve(path.dirname(path.join(repoRoot, importer)), specifier),
    );
    return absolute ? normalize(path.relative(repoRoot, absolute)) : undefined;
  }
  for (const [name, workspace] of workspaces) {
    if (specifier !== name && !specifier.startsWith(`${name}/`)) continue;
    const target = exportedTarget(workspace, specifier.slice(name.length + 1));
    if (!target) return undefined;
    const absolute = resolveFile(path.join(repoRoot, workspace.directory, target));
    return absolute ? normalize(path.relative(repoRoot, absolute)) : undefined;
  }
  if (
    specifier.startsWith("node:") ||
    Bun.resolveSync(specifier, path.dirname(path.join(repoRoot, importer)))
  ) {
    return null;
  }
};

const hasNonLiteralDynamicImport = (source: string) => {
  if (/\bimport\s*\(\s*`(?:\\.|[^`\\])*\$\{/u.test(source)) return true;
  const masked = source.replace(
    /("(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|`(?:\\.|[^`\\])*`)|\/\*[\s\S]*?\*\/|\/\/[^\n]*/gu,
    (match, literal: string | undefined) => literal ?? " ".repeat(match.length),
  );
  return /\bimport\s*\(\s*(?!["'`])/u.test(masked);
};

const transpilerLoader = (file: string) => {
  const extension = path.extname(file).slice(1).replace("m", "").replace("c", "");
  if (extension === "tsx" || extension === "jsx") return extension;
  return extension === "js" ? "js" : "ts";
};

/** Type-only imports are omitted by scanImports; they cannot affect runtime selection. */
export const buildModuleGraph = (
  repoRoot = REPO_ROOT,
  additionalFiles: Iterable<string> = [],
): ModuleGraph => {
  const codeFiles = new Set<string>();
  for (const root of GRAPH_ROOTS) {
    for (const file of new Bun.Glob(`${root}/**/*`).scanSync({ cwd: repoRoot, onlyFiles: true })) {
      if (CODE_FILE.test(file) && !/\.d\.[cm]?ts$/u.test(file)) codeFiles.add(normalize(file));
    }
  }
  for (const file of additionalFiles) {
    if (CODE_FILE.test(file) && existsSync(path.join(repoRoot, file)))
      codeFiles.add(normalize(file));
  }
  const reverse = new Map<string, Set<string>>();
  const unresolved = new Map<string, string[]>();
  const workspaces = readWorkspaces(repoRoot);
  const transpilers = new Map<string, Bun.Transpiler>();
  const pending = [...codeFiles];
  const scanned = new Set<string>();
  while (pending.length > 0) {
    const importer = pending.shift();
    if (!importer || scanned.has(importer)) continue;
    scanned.add(importer);
    const source = readFileSync(path.join(repoRoot, importer), "utf8").replace(/^#![^\n]*\n/u, "");
    const loader = transpilerLoader(importer);
    let transpiler = transpilers.get(loader);
    if (!transpiler) {
      transpiler = new Bun.Transpiler({ loader });
      transpilers.set(loader, transpiler);
    }
    const failures: string[] = [];
    // scanImports intentionally omits non-literal imports, so detect that unsafe form separately.
    if (hasNonLiteralDynamicImport(source)) failures.push("non-literal dynamic import");
    let imports = transpiler.scanImports("");
    try {
      imports = transpiler.scanImports(source);
    } catch {
      failures.push("source could not be scanned");
    }
    for (const { path: specifier } of imports) {
      let resolved: string | null | undefined;
      try {
        resolved = resolveSpecifier(specifier, importer, repoRoot, workspaces);
      } catch {
        resolved = undefined;
      }
      if (resolved === undefined) failures.push(specifier);
      if (resolved === null || resolved === undefined) continue;
      const importers = reverse.get(resolved) ?? new Set<string>();
      importers.add(importer);
      reverse.set(resolved, importers);
      if (
        CODE_FILE.test(resolved) &&
        !/\.d\.[cm]?ts$/u.test(resolved) &&
        !resolved.startsWith("../") &&
        !resolved.split("/").includes("node_modules") &&
        !codeFiles.has(resolved)
      ) {
        codeFiles.add(resolved);
        pending.push(resolved);
      }
    }
    if (failures.length > 0) unresolved.set(importer, failures);
  }
  return { reverse, unresolved, codeFiles };
};

export type FastSelection = {
  suites: TestSuite[];
  reason: "closure" | "package-data" | "full-fallback";
  cause?: string;
  selectedCount: number;
  fullCount: number;
};

/** Select runtime dependants while retaining each owning command and its preloads. */
export const selectFastTestSuitesWithReason = ({
  suites,
  changedFiles,
  focusedFiles = focusedTests,
  repoRoot = REPO_ROOT,
  graph: suppliedGraph,
}: FastTestSuitesOptions): FastSelection => {
  const changed = new Set(changedFiles);
  const discovered = new Set(suites.flatMap(({ files }) => files));
  const focused = new Set(
    focusedFiles.filter((file) => discovered.has(file) || existsSync(path.join(repoRoot, file))),
  );
  const fullCount = new Set([...discovered, ...focused]).size;
  const preloadFiles = suites.flatMap(({ cwd, preloads }) =>
    preloads.map((preload) =>
      normalize(path.relative(repoRoot, path.resolve(repoRoot, cwd, preload))),
    ),
  );
  const graph = suppliedGraph ?? buildModuleGraph(repoRoot, preloadFiles);
  const packageData = new Set<string>();
  for (const file of changedFiles) {
    const match = /^packages\/([^/]+)\/(.+)$/u.exec(file);
    if (match && !CODE_FILE.test(file)) packageData.add(match[1]);
  }
  let fallbackCause: string | undefined;
  for (const file of changedFiles) {
    const isUntrackedCode = CODE_FILE.test(file) && !/\.d\.[cm]?ts$/u.test(file);
    if ((SOURCE_FILE.test(file) || isUntrackedCode) && !graph.codeFiles.has(file)) {
      fallbackCause = `changed source is absent from graph: ${file}`;
      break;
    }
  }
  const unknownDependencies = changedFiles.some((file) => CODE_FILE.test(file))
    ? [...graph.unresolved.keys()].toSorted()
    : [];
  const unknownCause =
    unknownDependencies.length > 0
      ? `unknown dependencies in ${unknownDependencies.length} ${unknownDependencies.length === 1 ? "file" : "files"}: ${unknownDependencies.slice(0, 3).join(", ")}`
      : undefined;
  const reached = new Set([...changedFiles, ...unknownDependencies]);
  const queue = [...reached];
  while (queue.length > 0) {
    const file = queue.shift();
    if (!file) continue;
    for (const importer of graph.reverse.get(file) ?? []) {
      if (reached.has(importer)) continue;
      reached.add(importer);
      queue.push(importer);
    }
  }
  const candidates = new Set<string>(focused);
  const suitesReachedThroughPreload = new Set(
    suites.filter(({ cwd, preloads }) =>
      preloads.some((preload) =>
        reached.has(normalize(path.relative(repoRoot, path.resolve(repoRoot, cwd, preload)))),
      ),
    ),
  );
  for (const suite of suitesReachedThroughPreload) {
    for (const file of suite.files) candidates.add(file);
  }
  for (const file of discovered) {
    if (changed.has(file) || reached.has(file)) candidates.add(file);
    const packageName = /^packages\/([^/]+)\//u.exec(file)?.at(1);
    if (packageName && packageData.has(packageName)) candidates.add(file);
  }
  if (!fallbackCause && fullCount > 0 && candidates.size / fullCount >= FULL_SELECTION_RATIO) {
    fallbackCause = `selection reached ${candidates.size}/${fullCount} files`;
  }
  if (fallbackCause) {
    return {
      suites: selectFullTestSuites({ suites, focusedFiles }),
      reason: "full-fallback",
      cause: unknownCause ? `${unknownCause}; ${fallbackCause}` : fallbackCause,
      selectedCount: fullCount,
      fullCount,
    };
  }
  const selected = suites.map((suite) => {
    return {
      ...suite,
      files: suite.files.filter((file) => candidates.has(file) && !focused.has(file)),
    };
  });
  selected.push({ cwd: ".", preloads: [], files: [...focused] });
  const reason = packageData.size > 0 ? "package-data" : "closure";
  return {
    suites: selected.filter(({ files }) => files.length > 0),
    reason,
    selectedCount: candidates.size,
    fullCount,
    cause: unknownCause,
  };
};

export const selectFastTestSuites = (options: FastTestSuitesOptions): TestSuite[] =>
  selectFastTestSuitesWithReason(options).suites;

if (import.meta.main) {
  const args = process.argv.slice(2);
  const shard = Number(args.at(args.indexOf("--shard") + 1));
  const depth = args.at(args.indexOf("--depth") + 1);
  if (
    !Number.isInteger(shard) ||
    shard < 1 ||
    shard > TEST_SHARD_COUNT ||
    (depth !== "fast" && depth !== "full")
  ) {
    panic("Usage: bun scripts/ci-test-shards.ts --shard <1..4> --depth <fast|full>");
  }
  const suites = discoverTestSuites();
  const allFiles = [...suites.flatMap(({ files }) => files), ...focusedTests];
  const assigned = assignTestShards(allFiles).at(shard - 1);
  if (!assigned) panic("Test shard unavailable");
  const base = process.env["CI_TEST_BASE"];
  const changedFiles = base ? changedTestPaths(base) : [];
  const fastSelection =
    depth === "fast" ? selectFastTestSuitesWithReason({ suites, changedFiles }) : undefined;
  const selectedSuites = fastSelection?.suites ?? selectFullTestSuites({ suites });
  const selectedFileCount = new Set(selectedSuites.flatMap(({ files }) => files)).size;
  console.log(
    `Shard ${shard}/${TEST_SHARD_COUNT}: changed=${changedFiles.length} selected=${selectedFileCount} reason=${fastSelection?.reason ?? "full"}${fastSelection?.cause ? ` (${fastSelection.cause})` : ""}`,
  );
  let exitCode = 0;
  for (const { cwd, preloads, files } of selectedSuites) {
    const selected = files.filter((file) => assigned.has(file));
    if (selected.length === 0) continue;
    console.log(`Shard ${shard}/${TEST_SHARD_COUNT}: ${cwd} (${selected.length} files)`);
    const result = Bun.spawnSync(
      [
        "bun",
        "test",
        ...preloads.flatMap((preload) => ["--preload", preload]),
        ...selected.map((file) => `./${path.posix.relative(cwd, file)}`),
      ],
      { cwd: path.join(REPO_ROOT, cwd), stdout: "inherit", stderr: "inherit" },
    );
    if (result.exitCode !== 0) exitCode = 1;
  }
  process.exit(exitCode);
}
