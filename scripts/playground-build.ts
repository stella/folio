import { createHash } from "node:crypto";
import {
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from "node:fs";
import { execFileSync, spawnSync } from "node:child_process";
import { join, relative, resolve, sep } from "node:path";
import { Result, TaggedError } from "better-result";

const CACHE_FILE = ".playground-build-cache.json";
const CACHE_VERSION = 1;
const VITE_ENV_FILES = [".env", ".env.local", ".env.production", ".env.production.local"] as const;

type FileMetadata = { path: string; size: number; mtimeMs: number };
type CacheRecord = {
  version: typeof CACHE_VERSION;
  inputFingerprint: string;
  outputFingerprint: string;
};

class PlaygroundBuildError extends TaggedError("PlaygroundBuildError")<{
  message: string;
  cause?: unknown;
}> {}

type PlaygroundBuildResult = {
  status: "built" | "skipped";
  elapsedMs: number;
};

type EnsurePlaygroundBuildOptions = {
  repoRoot: string;
  packageDir: string;
  /** Test seam; production invokes `bun --filter <package-name> build`. */
  build?: () => Promise<void> | void;
  /** Test seam for build variables; production fingerprints the current environment. */
  buildEnvironment?: Record<string, string | undefined>;
  log?: (message: string) => void;
  now?: () => number;
};

const normalizedPath = (path: string): string => path.split(sep).join("/");

const buildEnvironmentFingerprint = (
  environment: Record<string, string | undefined>,
): Record<string, string> => ({
  bunVersion: Bun.version,
  ...Object.fromEntries(
    Object.entries(environment)
      .filter(
        ([name, value]) => value !== undefined && (name === "NODE_ENV" || name.startsWith("VITE_")),
      )
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([name, value]) => [name, value ?? ""]),
  ),
});

const fingerprint = (
  metadata: readonly FileMetadata[],
  environment: Record<string, string> = {},
): string => createHash("sha256").update(JSON.stringify({ metadata, environment })).digest("hex");

const packageInputs = (repoRoot: string, packageDir: string): FileMetadata[] => {
  const listed = execFileSync("git", ["ls-files", "-co", "--exclude-standard", "-z"], {
    cwd: repoRoot,
    encoding: "utf8",
  })
    .split("\0")
    .filter(Boolean);
  const paths = new Set(listed);
  for (const envFile of VITE_ENV_FILES) {
    const packageEnvFile = relative(repoRoot, join(packageDir, envFile));
    if (existsSync(join(repoRoot, packageEnvFile))) paths.add(packageEnvFile);
  }

  return [...paths].sort().flatMap((path) => {
    const packageRelativePath = relative(packageDir, join(repoRoot, path));
    if (packageRelativePath === "dist" || packageRelativePath.startsWith(`dist${sep}`)) return [];
    const absolutePath = join(repoRoot, path);
    if (!existsSync(absolutePath)) return [];
    const stat = statSync(absolutePath);
    if (!stat.isFile()) return [];
    return [{ path: normalizedPath(path), size: stat.size, mtimeMs: stat.mtimeMs }];
  });
};

const outputFiles = (distDir: string): FileMetadata[] => {
  if (!existsSync(distDir)) return [];
  const files: FileMetadata[] = [];
  const visit = (directory: string): void => {
    for (const entry of readdirSync(directory, { withFileTypes: true })) {
      const absolutePath = join(directory, entry.name);
      if (entry.isDirectory()) {
        visit(absolutePath);
        continue;
      }
      if (!entry.isFile() || entry.name === CACHE_FILE) continue;
      const stat = statSync(absolutePath);
      files.push({
        path: normalizedPath(relative(distDir, absolutePath)),
        size: stat.size,
        mtimeMs: stat.mtimeMs,
      });
    }
  };
  visit(distDir);
  return files.sort((left, right) => left.path.localeCompare(right.path));
};

const isCacheRecord = (value: unknown): value is CacheRecord =>
  typeof value === "object" &&
  value !== null &&
  "version" in value &&
  value.version === CACHE_VERSION &&
  "inputFingerprint" in value &&
  typeof value.inputFingerprint === "string" &&
  "outputFingerprint" in value &&
  typeof value.outputFingerprint === "string";

const readCache = (cachePath: string): CacheRecord | null => {
  if (!existsSync(cachePath)) return null;
  const parsed = Result.try({
    try: (): unknown => JSON.parse(readFileSync(cachePath, "utf8")),
    catch: () => null,
  });
  if (parsed.isErr() || !isCacheRecord(parsed.value)) return null;
  return parsed.value;
};

const runBuild = (repoRoot: string, packageDir: string): void => {
  const manifest: unknown = JSON.parse(readFileSync(join(packageDir, "package.json"), "utf8"));
  const name =
    typeof manifest === "object" && manifest !== null && "name" in manifest
      ? manifest.name
      : undefined;
  if (typeof name !== "string" || name.length === 0) {
    throw new TypeError(`Package manifest at ${packageDir} has no package name`);
  }
  const result = spawnSync("bun", ["--filter", name, "build"], {
    cwd: repoRoot,
    stdio: "inherit",
  });
  if (result.error) {
    throw new PlaygroundBuildError({
      message: `Playground build failed for ${name}`,
      cause: result.error,
    });
  }
  if (result.status !== 0) {
    throw new PlaygroundBuildError({ message: `Playground build failed for ${name}` });
  }
};

export const ensurePlaygroundBuild = async ({
  repoRoot,
  packageDir,
  build,
  buildEnvironment = process.env,
  log = console.log,
  now = Date.now,
}: EnsurePlaygroundBuildOptions): Promise<PlaygroundBuildResult> => {
  const startedAt = now();
  const root = resolve(repoRoot);
  const packagePath = resolve(root, packageDir);
  const distDir = join(packagePath, "dist");
  const cachePath = join(distDir, CACHE_FILE);
  const inputsBeforeBuild = packageInputs(root, packagePath);
  const environment = buildEnvironmentFingerprint(buildEnvironment);
  const inputFingerprint = fingerprint(inputsBeforeBuild, environment);
  const currentCache = readCache(cachePath);
  const currentOutputs = outputFiles(distDir);

  if (
    currentCache?.version === CACHE_VERSION &&
    currentCache.inputFingerprint === inputFingerprint &&
    currentCache.outputFingerprint === fingerprint(currentOutputs) &&
    currentOutputs.some(({ path }) => path === "index.html")
  ) {
    const elapsedMs = now() - startedAt;
    log(`Playground build skipped (cache hit, ${elapsedMs} ms)`);
    return { status: "skipped", elapsedMs };
  }

  if (existsSync(cachePath)) unlinkSync(cachePath);
  await (build ?? (() => runBuild(root, packagePath)))();
  const inputsAfterBuild = packageInputs(root, packagePath);
  const environmentAfterBuild = buildEnvironmentFingerprint(buildEnvironment);
  if (fingerprint(inputsAfterBuild, environmentAfterBuild) !== inputFingerprint) {
    throw new PlaygroundBuildError({
      message: "Playground inputs or build environment changed during the build",
    });
  }

  const outputs = outputFiles(distDir);
  if (!outputs.some(({ path }) => path === "index.html")) {
    throw new PlaygroundBuildError({
      message: `Playground build produced no index.html in ${distDir}`,
    });
  }
  mkdirSync(distDir, { recursive: true });
  const record: CacheRecord = {
    version: CACHE_VERSION,
    inputFingerprint,
    outputFingerprint: fingerprint(outputs),
  };
  writeFileSync(cachePath, `${JSON.stringify(record)}\n`);
  const elapsedMs = now() - startedAt;
  log(`Playground build completed and cached (${elapsedMs} ms)`);
  return { status: "built", elapsedMs };
};

const main = async (): Promise<void> => {
  const packageArgument = Bun.argv[2];
  if (packageArgument === undefined) {
    throw new TypeError("Usage: bun scripts/playground-build.ts <playground package directory>");
  }
  const repoRoot = resolve(import.meta.dir, "..");
  await ensurePlaygroundBuild({ repoRoot, packageDir: packageArgument });
};

if (import.meta.main) await main();
