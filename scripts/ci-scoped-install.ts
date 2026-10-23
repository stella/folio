#!/usr/bin/env bun
/** Install a task environment from the repository lock, then restore manifests. */
import {
  existsSync,
  readdirSync,
  readFileSync,
  realpathSync,
  statSync,
  writeFileSync,
} from "node:fs";
import path from "node:path";

import { installFrozenRepositoryDependencies } from "./clean-room-install";

export const INSTALL_SCOPES = {
  "ci-plan": {
    roots: ["@stll/oxlint-config", "oxlint", "oxlint-tsgolint"],
    workspaces: "none",
  },
  "browser-discovery": {
    roots: [
      "@playwright/test",
      "typescript",
      "fast-check",
      "better-result",
      "jszip",
      "@stll/folio-core",
      "prosemirror-state",
    ],
    workspaces: "runtime",
  },
  "typecheck-budget": {
    roots: [
      "@typescript/native",
      "typescript",
      "better-result",
      "bun-types",
      "@types/node",
      "@types/react",
      "@types/react-dom",
      "fast-check",
      "jszip",
      "pixelmatch",
      "pngjs",
      "@types/pngjs",
      "@playwright/test",
      "@stll/folio-core",
      "prosemirror-state",
    ],
    workspaces: "all",
  },
} as const;

export const isInstallScope = (
  scope: string,
): scope is keyof typeof INSTALL_SCOPES => Object.hasOwn(INSTALL_SCOPES, scope);

type Manifest = {
  name: string;
  workspaces?: string[];
  dependencies?: Record<string, string>;
  devDependencies?: Record<string, string>;
  optionalDependencies?: Record<string, string>;
  peerDependencies?: Record<string, string>;
};
export type Lock = {
  workspaces: Record<string, Manifest>;
  packages: Record<string, [string, ...unknown[]]>;
};

/** Logical package bytes, counting shared files once and excluding workspace sources. */
export const installedPackageBytes = (repoRoot: string): number => {
  const physicalRoot = realpathSync(repoRoot);
  const directories = new Set<string>();
  const files = new Set<string>();
  const pending = [path.join(physicalRoot, "node_modules")];
  let bytes = 0;
  while (pending.length > 0) {
    const item = pending.pop();
    if (item === undefined) break;
    const resolved = realpathSync(item);
    if (resolved.startsWith(path.join(physicalRoot, "packages") + path.sep))
      continue;
    const stat = statSync(resolved);
    if (stat.isDirectory()) {
      if (directories.has(resolved)) continue;
      directories.add(resolved);
      for (const name of readdirSync(resolved))
        pending.push(path.join(resolved, name));
      continue;
    }
    const identity = `${stat.dev}:${stat.ino}`;
    if (files.has(identity)) continue;
    files.add(identity);
    bytes += stat.size;
  }
  return bytes;
};

/**
 * Bun 1.4.4 drops lockfile packages a scoped install no longer reaches, so a
 * pruned lock is not frozen-clean by itself. Bun writes its own pruned lock
 * here, and it may only remove packages or re-hoist an already locked one (a
 * nested copy becoming the top-level entry): every entry it keeps must equal
 * an entry of the pruned lock, and any other changed field fails, so a scoped
 * install still resolves nothing new.
 */
/** JSON with object keys sorted, so key order never reads as a lockfile change. */
const canonical = (value: unknown): string =>
  JSON.stringify(value, (_, inner: unknown) =>
    inner !== null && typeof inner === "object" && !Array.isArray(inner)
      ? Object.fromEntries(
          Object.entries(inner).toSorted(([left], [right]) =>
            left < right ? -1 : left > right ? 1 : 0,
          ),
        )
      : inner,
  );

/** Why a settled lock is not a subset of the pruned one, or nothing. */
export const prunedLockChanges = (before: Lock, after: Lock): string[] => {
  const locked = new Set(
    Object.values(before.packages).map((entry) => canonical(entry)),
  );
  const resolved = Object.entries(after.packages)
    .filter(([, entry]) => !locked.has(canonical(entry)))
    .map(([name]) => `re-resolved ${name}`);
  const fields = [...new Set([...Object.keys(before), ...Object.keys(after)])]
    .filter(
      (key) =>
        key !== "packages" &&
        canonical(Reflect.get(before, key)) !==
          canonical(Reflect.get(after, key)),
    )
    .map((key) => `changed ${key}`);
  return [...resolved, ...fields];
};

export const settlePrunedLockfile = (repoRoot: string): void => {
  const file = path.join(repoRoot, "bun.lock");
  const before: Lock = Bun.JSONC.parse(readFileSync(file, "utf8"));
  const result = Bun.spawnSync(
    [process.execPath, "install", "--lockfile-only"],
    { cwd: repoRoot, stdout: "inherit", stderr: "inherit" },
  );
  if (result.exitCode !== 0)
    throw new TypeError("Scoped lockfile settle failed");
  const after: Lock = Bun.JSONC.parse(readFileSync(file, "utf8"));
  const changes = prunedLockChanges(before, after);
  if (changes.length > 0)
    throw new TypeError(`Scoped lockfile settle: ${changes.join("; ")}`);
};

export const scopeInstall = (
  scope: keyof typeof INSTALL_SCOPES,
  repoRoot: string,
) => {
  const originals = new Map<string, string>();
  const read = (file: string) => {
    const text = readFileSync(path.join(repoRoot, file), "utf8");
    originals.set(file, text);
    return text;
  };
  const root: Manifest = JSON.parse(read("package.json"));
  const lock: Lock = Bun.JSONC.parse(read("bun.lock"));
  const profile = INSTALL_SCOPES[scope];
  const dependencies: Record<string, string> = {};
  for (const name of profile.roots) {
    const version = root.devDependencies?.[name];
    if (version === undefined)
      throw new TypeError(`Undeclared scoped dependency: ${name}`);
    dependencies[name] = version;
  }
  root.devDependencies = dependencies;
  const rootLock = lock.workspaces[""];
  if (rootLock === undefined)
    throw new TypeError("Lockfile has no root workspace");
  rootLock.devDependencies = dependencies;

  const retained = new Set<string>();
  const workspaceByName = new Map(
    Object.entries(lock.workspaces).map(([dir, manifest]) => [
      manifest.name,
      dir,
    ]),
  );
  const retainRuntime = (name: string) => {
    const dir = workspaceByName.get(name);
    if (dir === undefined) throw new TypeError(`Unknown workspace: ${name}`);
    if (retained.has(dir)) return;
    retained.add(dir);
    const manifest = lock.workspaces[dir];
    if (manifest === undefined)
      throw new TypeError(`Missing workspace: ${dir}`);
    for (const [dependency, version] of Object.entries(
      manifest.dependencies ?? {},
    )) {
      if (version.startsWith("workspace:")) retainRuntime(dependency);
    }
  };
  if (profile.workspaces === "runtime") retainRuntime("@stll/folio-cli");
  if (profile.workspaces === "all") {
    for (const dir of Object.keys(lock.workspaces))
      if (dir !== "") retained.add(dir);
  }
  root.workspaces = [...retained].toSorted();
  const files = new Map<string, string>();
  for (const [dir, manifest] of Object.entries(lock.workspaces)) {
    if (dir === "") continue;
    if (!retained.has(dir)) {
      continue;
    }
    if (profile.workspaces !== "runtime") continue;
    const file = `${dir}/package.json`;
    const source: Manifest = JSON.parse(read(file));
    delete source.devDependencies;
    delete manifest.devDependencies;
    files.set(file, JSON.stringify(source, null, 2));
  }
  lock.workspaces = Object.fromEntries(
    Object.entries(lock.workspaces).filter(
      ([dir]) => dir === "" || retained.has(dir),
    ),
  );
  lock.packages = Object.fromEntries(
    Object.entries(lock.packages).filter(([, entry]) => {
      const workspace = entry[0].split("@workspace:").at(1);
      return workspace === undefined || retained.has(workspace);
    }),
  );
  files.set("package.json", JSON.stringify(root, null, 2));
  files.set("bun.lock", JSON.stringify(lock, null, 2));
  return { files, originals };
};

if (import.meta.main) {
  const scope = process.argv.at(2);
  if (scope === undefined || !isInstallScope(scope))
    throw new TypeError("Expected a scoped CI job name");
  const repoRoot = path.resolve(import.meta.dir, "..");
  if (existsSync(path.join(repoRoot, "node_modules")))
    throw new TypeError("Scoped CI installs require a clean dependency tree");
  const run = (name: keyof typeof INSTALL_SCOPES) => {
    const { files, originals } = scopeInstall(name, repoRoot);
    const started = performance.now();
    try {
      for (const [file, text] of files)
        writeFileSync(path.join(repoRoot, file), text);
      settlePrunedLockfile(repoRoot);
      installFrozenRepositoryDependencies(repoRoot);
      if (existsSync(path.join(repoRoot, "node_modules/binaryen")))
        throw new TypeError(
          "Scoped install includes an unrelated WASM optimizer",
        );
      console.log(
        `Scoped install ${name}: ${Math.round(performance.now() - started)} ms`,
      );
      console.log(
        `Scoped install ${name}: ${installedPackageBytes(repoRoot)} package bytes`,
      );
    } finally {
      for (const [file, text] of originals)
        writeFileSync(path.join(repoRoot, file), text);
    }
  };
  run(scope);
}
