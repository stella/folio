#!/usr/bin/env bun
// Refresh ci-test-timings.json from per-file group spans in a successful full-suite log.
// Keep millisecond durations above one second; unmeasured files still run.
import { panic } from "better-result";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  statSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";

import {
  EXPLORATION_BASE_ENV,
  EXPLORATION_ENV,
  EXPLORATION_OWNED_ENV,
} from "../test/property-exploration";
import {
  PROPERTY_SEEDS_FILE,
  testFileForSeedFile,
} from "../test/seed-registry";
import focusedTests from "./ci-focused-tests.json";
import {
  discoverTestSuites,
  plannedTestFiles,
  planTestShards,
  shardLoads,
  type TestSuite,
  type TestTimings,
} from "./ci-test-shard-plan";
import { parseSourceModule, staticStringValue, walkAst } from "./lib/oxc-ast";

const REPO_ROOT = path.resolve(import.meta.dir, "..");

const decodeXml = (value: string) =>
  value
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&");

export const junitFileTimings = (xml: string, cwd: string): TestTimings => {
  const result: TestTimings = {};
  for (const match of xml.matchAll(/<testsuite\s+[^>]*>/gu)) {
    const tag = match[0];
    if (/\sline="/u.test(tag)) continue;
    const file = /\sfile="([^"]+)"/u.exec(tag)?.at(1);
    const seconds = /\stime="([0-9.]+)"/u.exec(tag)?.at(1);
    if (!file || !seconds) continue;
    const repoFile = path.posix.normalize(
      path.posix.join(cwd, decodeXml(file).replaceAll("\\", "/")),
    );
    result[repoFile] = Math.max(1, Math.round(Number(seconds) * 1_000));
  }
  return result;
};

export const focusedTestFiles = focusedTests;
const SOURCE_FILE = /^packages\/[^/]+\/src\/.*\.[cm]?[jt]sx?$/u;
const CODE_FILE = /\.[cm]?[jt]sx?$/u;
const TEST_FILE = /\.test\.[cm]?[jt]sx?$/u;
const CODE_EXTENSIONS = [
  "",
  ".ts",
  ".tsx",
  ".js",
  ".jsx",
  ".mts",
  ".cts",
  ".mjs",
  ".cjs",
];
const GRAPH_ROOTS = [
  "packages/*/src",
  "packages/*/test",
  "scripts",
  "benchmarks/compare",
  "test",
];
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
  const extra = [...new Set(focusedFiles)].filter(
    (file) => !discovered.has(file),
  );
  if (extra.length === 0) return suites;
  return [...suites, { cwd: ".", preloads: [], files: extra }];
};

/** Deleted paths never expand a fast suite; git returns rename destinations. */
export const changedTestPaths = (
  base: string,
  repoRoot = REPO_ROOT,
): string[] => {
  const result = Bun.spawnSync(
    [
      "git",
      "diff",
      "--name-only",
      "--diff-filter=ACMRT",
      "-z",
      `${base}...HEAD`,
      "--",
    ],
    { cwd: repoRoot, stdout: "pipe", stderr: "pipe" },
  );
  if (result.exitCode !== 0)
    panic(`Cannot select changed tests: ${result.stderr.toString()}`);
  return result.stdout
    .toString()
    .split("\0")
    .filter((file) => file !== "" && existsSync(path.join(repoRoot, file)));
};

/**
 * The changed paths, plus the test file each changed pinned seed file pins. A
 * seed file is read at runtime, not imported, so the module graph never
 * reaches the test it pins: changing the seeds counts as changing that test.
 */
export const changedTestFiles = (
  changedFiles: readonly string[],
): Set<string> =>
  new Set([
    ...changedFiles,
    ...changedFiles
      .filter((file) => file.startsWith(`${PROPERTY_SEEDS_FILE}/`))
      .map(testFileForSeedFile),
  ]);

/**
 * The test files a change adds or changes itself (directly or through its
 * pinned seeds). Their generated property failures fail a pull request or
 * merge group run, while every other file's still only report
 * (test/property-exploration.ts).
 */
export const explorationOwnedTestFiles = (
  changedFiles: readonly string[],
): string[] =>
  [...changedTestFiles(changedFiles)]
    .filter((file) => TEST_FILE.test(file))
    .toSorted();

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

type Workspace = {
  directory: string;
  exports: Record<string, string | null>;
  source?: string;
};

const readWorkspaces = (repoRoot: string) => {
  const workspaces = new Map<string, Workspace>();
  for (const manifest of new Bun.Glob("packages/*/package.json").scanSync({
    cwd: repoRoot,
  })) {
    const pkg = JSON.parse(readFileSync(path.join(repoRoot, manifest), "utf8"));
    if (typeof pkg.name !== "string") continue;
    workspaces.set(pkg.name, {
      directory: path.dirname(manifest),
      exports:
        typeof pkg.exports === "object" && pkg.exports ? pkg.exports : {},
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
    const absolute = resolveFile(
      path.join(repoRoot, workspace.directory, target),
    );
    return absolute ? normalize(path.relative(repoRoot, absolute)) : undefined;
  }
  if (specifier.startsWith("node:")) return null;
  const resolved = Bun.resolveSync(
    specifier,
    path.dirname(path.join(repoRoot, importer)),
  );
  const relative = normalize(path.relative(repoRoot, resolved));
  if (!relative.startsWith("../") && !path.isAbsolute(relative)) {
    if (!relative.split("/").includes("node_modules")) return relative;
  }
  return null;
};

const unsafeModuleLoads = (source: string, file: string) => {
  const { program } = parseSourceModule(file, source);
  const failures = new Set<string>();
  walkAst(program, (node, ancestors) => {
    if (
      node.type === "ImportExpression" &&
      (node.options !== null || staticStringValue(node.source) === undefined)
    ) {
      failures.add("non-literal dynamic import");
    }
    // JSX names are identifiers to TypeScript too.
    const identifier =
      node.type === "Identifier" || node.type === "JSXIdentifier"
        ? node.name
        : undefined;
    if (identifier === "require") {
      const parent = ancestors.at(-1);
      const argument =
        parent?.type === "CallExpression" ? parent.arguments.at(0) : undefined;
      if (
        parent?.type !== "CallExpression" ||
        parent.callee !== node ||
        parent.arguments.length !== 1 ||
        argument === undefined ||
        argument.type !== "Literal" ||
        typeof argument.value !== "string"
      ) {
        failures.add("unsafe require reference");
      }
    }
    if (identifier === "createRequire") {
      failures.add("createRequire reference");
    }
    if (
      node.type === "MemberExpression" &&
      !node.computed &&
      node.property.type === "Identifier"
    ) {
      if (
        node.object.type === "Identifier" &&
        node.object.name === "module" &&
        node.property.name === "require"
      ) {
        failures.add("module.require reference");
      }
      if (
        node.property.name === "resolve" &&
        node.object.type === "MetaProperty" &&
        node.object.meta.name === "import"
      ) {
        failures.add("import.meta.resolve reference");
      }
    }
    if (
      node.type === "NewExpression" &&
      node.callee.type === "Identifier" &&
      node.callee.name === "Worker"
    ) {
      failures.add("Worker constructor");
    }
  });
  return [...failures];
};

const transpilerLoader = (file: string) => {
  const extension = path
    .extname(file)
    .slice(1)
    .replace("m", "")
    .replace("c", "");
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
    for (const file of new Bun.Glob(`${root}/**/*`).scanSync({
      cwd: repoRoot,
      onlyFiles: true,
    })) {
      if (CODE_FILE.test(file) && !/\.d\.[cm]?ts$/u.test(file))
        codeFiles.add(normalize(file));
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
    const source = readFileSync(path.join(repoRoot, importer), "utf8").replace(
      /^#![^\n]*\n/u,
      "",
    );
    const loader = transpilerLoader(importer);
    let transpiler = transpilers.get(loader);
    if (!transpiler) {
      transpiler = new Bun.Transpiler({ loader });
      transpilers.set(loader, transpiler);
    }
    const failures: string[] = [];
    let imports = transpiler.scanImports("");
    try {
      failures.push(...unsafeModuleLoads(source, importer));
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
  const changed = changedTestFiles(changedFiles);
  const discovered = new Set(suites.flatMap(({ files }) => files));
  const focused = new Set(
    focusedFiles.filter(
      (file) => discovered.has(file) || existsSync(path.join(repoRoot, file)),
    ),
  );
  const fullCount = new Set([...discovered, ...focused]).size;
  // Resolve each preload the way the suite command loads it, so an
  // extensionless or directory preload still joins the graph. A supplied graph
  // may name the preload path directly.
  const workspaces = readWorkspaces(repoRoot);
  const unresolvedPreloads: string[] = [];
  const suitePreloads = new Map(
    suites.map((suite) => [
      suite,
      suite.preloads.flatMap((preload) => {
        let resolved: string | null | undefined;
        try {
          resolved = resolveSpecifier(
            preload,
            path.posix.join(suite.cwd, "package.json"),
            repoRoot,
            workspaces,
          );
        } catch {
          resolved = undefined;
        }
        const literal = path.posix.join(suite.cwd, preload);
        if (resolved === undefined && suppliedGraph?.codeFiles.has(literal))
          resolved = literal;
        if (resolved === undefined) unresolvedPreloads.push(literal);
        return resolved ? [resolved] : [];
      }),
    ]),
  );
  const preloadFiles = [...suitePreloads.values()].flat();
  const graph = suppliedGraph ?? buildModuleGraph(repoRoot, preloadFiles);
  const packageData = new Set<string>();
  for (const file of changedFiles) {
    const match = /^packages\/([^/]+)\/(.+)$/u.exec(file);
    if (match && !CODE_FILE.test(file)) packageData.add(match[1]);
  }
  let fallbackCause: string | undefined =
    unresolvedPreloads.length > 0
      ? `suite preload cannot be resolved: ${unresolvedPreloads.join(", ")}`
      : undefined;
  for (const file of fallbackCause ? [] : changedFiles) {
    const isUntrackedCode =
      CODE_FILE.test(file) && !/\.d\.[cm]?ts$/u.test(file);
    if (
      (SOURCE_FILE.test(file) || isUntrackedCode) &&
      !graph.codeFiles.has(file)
    ) {
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
    suites.filter((suite) =>
      (suitePreloads.get(suite) ?? []).some((preload) => reached.has(preload)),
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
  if (
    !fallbackCause &&
    fullCount > 0 &&
    candidates.size / fullCount >= FULL_SELECTION_RATIO
  ) {
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
      files: suite.files.filter(
        (file) => candidates.has(file) && !focused.has(file),
      ),
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

/**
 * The environment the shard's test processes run with. When generated
 * property failures only report (pull requests and merge groups), it names
 * the test files the change owns in EXPLORATION_OWNED_ENV, from the diff
 * against EXPLORATION_BASE_ENV; a report run without that base fails, since
 * it could not tell a change's own properties from everyone else's.
 */
export const explorationTestEnv = (
  env: NodeJS.ProcessEnv,
  changedPaths: (base: string) => string[] = changedTestPaths,
): NodeJS.ProcessEnv => {
  if (env[EXPLORATION_ENV] !== "report") return env;
  const base = env[EXPLORATION_BASE_ENV];
  if (base === undefined || base === "")
    return panic(
      `${EXPLORATION_ENV}=report needs ${EXPLORATION_BASE_ENV}, the commit the change is compared with`,
    );
  const owned = explorationOwnedTestFiles(changedPaths(base));
  console.log(
    `Exploration: ${String(owned.length)} test files this change adds or changes fail on generated cases${owned.length > 0 ? `: ${owned.join(", ")}` : ""}`,
  );
  return { ...env, [EXPLORATION_OWNED_ENV]: JSON.stringify(owned) };
};

export const selectFastTestSuites = (
  options: FastTestSuitesOptions,
): TestSuite[] => selectFastTestSuitesWithReason(options).suites;

if (import.meta.main) {
  const args = process.argv.slice(2);
  const shard = Number(args.at(args.indexOf("--shard") + 1));
  const total = Number(args.at(args.indexOf("--shards") + 1));
  const depth = args.at(args.indexOf("--depth") + 1);
  if (
    !Number.isInteger(shard) ||
    !Number.isInteger(total) ||
    shard < 1 ||
    shard > total ||
    (depth !== "fast" && depth !== "full")
  ) {
    panic(
      "Usage: bun scripts/ci-test-shards.ts --shard <k> --shards <n> --depth <fast|full>",
    );
  }
  const suites = discoverTestSuites();
  const shards = planTestShards(plannedTestFiles());
  // The workflow matrix comes from the same plan; a different count means the
  // job runs against another commit's measurements.
  if (shards.length !== total)
    panic(`The plan has ${shards.length} test shards, the job runs ${total}`);
  const loads = shardLoads(shards);
  console.log(
    `Predicted shard loads: ${loads.map((load, index) => `${index + 1}=${load}ms`).join(", ")}`,
  );
  const assigned = shards.at(shard - 1);
  if (!assigned) panic("Test shard unavailable");
  const base = process.env["CI_TEST_BASE"];
  const changedFiles = base ? changedTestPaths(base) : [];
  const testEnv = explorationTestEnv(process.env);
  const fastSelection =
    depth === "fast"
      ? selectFastTestSuitesWithReason({ suites, changedFiles })
      : undefined;
  const selectedSuites =
    fastSelection?.suites ?? selectFullTestSuites({ suites });
  const selectedFileCount = new Set(
    selectedSuites.flatMap(({ files }) => files),
  ).size;
  console.log(
    `Shard ${shard}/${total}: changed=${changedFiles.length} selected=${selectedFileCount} reason=${fastSelection?.reason ?? "full"}${fastSelection?.cause ? ` (${fastSelection.cause})` : ""}`,
  );
  let exitCode = 0;
  const recorded: TestTimings = {};
  const resultsDirectory = path.join(REPO_ROOT, "test-results");
  if (depth === "full") mkdirSync(resultsDirectory, { recursive: true });
  for (const [
    suiteIndex,
    { cwd, preloads, files },
  ] of selectedSuites.entries()) {
    const selected = files.filter((file) => assigned.has(file));
    if (selected.length === 0) continue;
    console.log(`Shard ${shard}/${total}: ${cwd} (${selected.length} files)`);
    const junitFile = path.join(
      resultsDirectory,
      `junit-${shard}-${suiteIndex}.xml`,
    );
    const result = Bun.spawnSync(
      [
        "bun",
        "test",
        ...(depth === "full"
          ? ["--reporter=junit", `--reporter-outfile=${junitFile}`]
          : []),
        ...preloads.flatMap((preload) => ["--preload", preload]),
        ...selected.map((file) => `./${path.posix.relative(cwd, file)}`),
      ],
      {
        cwd: path.join(REPO_ROOT, cwd),
        env: testEnv,
        stdout: "inherit",
        stderr: "inherit",
      },
    );
    if (depth === "full" && existsSync(junitFile)) {
      Object.assign(
        recorded,
        junitFileTimings(readFileSync(junitFile, "utf8"), cwd),
      );
    }
    if (result.exitCode !== 0) exitCode = 1;
  }
  if (depth === "full") {
    writeFileSync(
      path.join(resultsDirectory, `test-timings-${shard}.json`),
      `${JSON.stringify({ runId: process.env["GITHUB_RUN_ID"] ?? "local", timings: recorded }, null, 2)}\n`,
    );
  }
  process.exit(exitCode);
}
