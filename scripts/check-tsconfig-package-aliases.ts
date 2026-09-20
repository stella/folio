#!/usr/bin/env bun

import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";

const ROOT = path.join(import.meta.dirname, "..");
const PACKAGES_DIRECTORY = path.join(ROOT, "packages");
const TSCONFIG_DEPCRUISE_PATH = path.join(ROOT, "tsconfig.depcruise.json");
const PACKAGE_ALIAS_RE = /^(@stll\/[^/]+)(?:\/.*)?$/;

type WorkspacePackage = { directory: string; name: string };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const parseJson = (filePath: string): unknown => JSON.parse(readFileSync(filePath, "utf8"));

const readWorkspacePackages = (): WorkspacePackage[] =>
  readdirSync(PACKAGES_DIRECTORY, { withFileTypes: true }).flatMap((entry) => {
    const manifestPath = path.join(PACKAGES_DIRECTORY, entry.name, "package.json");
    if (!entry.isDirectory() || !existsSync(manifestPath)) return [];
    const manifest = parseJson(manifestPath);
    if (!isRecord(manifest) || typeof manifest["name"] !== "string") return [];
    return [{ directory: entry.name, name: manifest["name"] }];
  });

const readTSConfigPaths = (): Record<string, readonly string[]> => {
  const tsconfig = parseJson(TSCONFIG_DEPCRUISE_PATH);
  const paths =
    isRecord(tsconfig) && isRecord(tsconfig["compilerOptions"])
      ? tsconfig["compilerOptions"]["paths"]
      : undefined;
  if (!isRecord(paths)) return {};
  return Object.fromEntries(
    Object.entries(paths).flatMap(([key, value]) =>
      Array.isArray(value) && value.every((entry) => typeof entry === "string")
        ? [[key, value as readonly string[]]]
        : [],
    ),
  );
};

/**
 * The no-circular guard's `*-uses-package-contracts` rules exempt the
 * "aliased-tsconfig-paths" dependency type from `dependencyTypesNot`,
 * treating any import dependency-cruiser resolved through a tsconfig `paths`
 * entry as equivalent to importing the target's declared package name.
 * dependency-cruiser tags that way purely because the specifier matched a
 * `paths` key, with no check on the key's content; the equivalence only
 * holds if every `paths` entry actually is a workspace package name pointing
 * into that same package's own src. This validates that invariant so a
 * stray or mistargeted alias fails loudly here instead of silently opening a
 * hole in the boundary rule.
 */
export const validateTSConfigPackageAliases = (
  packages: readonly WorkspacePackage[],
  paths: Record<string, readonly string[]>,
): string[] => {
  const directoryByPackageName = new Map(packages.map(({ directory, name }) => [name, directory]));
  const issues: string[] = [];

  for (const [key, targets] of Object.entries(paths)) {
    const match = PACKAGE_ALIAS_RE.exec(key);
    if (!match) {
      issues.push(
        `tsconfig.depcruise.json paths entry "${key}" is not a @stll/<package> alias: the ` +
          "no-circular guard treats every tsconfig path alias as equivalent to the target's " +
          "package name, so a non-package alias must not exist",
      );
      continue;
    }
    const packageName = match[1] as string;
    const directory = directoryByPackageName.get(packageName);
    if (directory === undefined) {
      issues.push(
        `tsconfig.depcruise.json paths entry "${key}" references unknown package "${packageName}"`,
      );
      continue;
    }
    const expectedPrefix = `packages/${directory}/src/`;
    for (const target of targets) {
      if (!target.startsWith(expectedPrefix)) {
        issues.push(
          `tsconfig.depcruise.json paths entry "${key}" -> "${target}" must stay under ${expectedPrefix}`,
        );
      }
    }
  }
  return issues.toSorted();
};

const main = () => {
  const issues = validateTSConfigPackageAliases(readWorkspacePackages(), readTSConfigPaths());
  if (issues.length === 0) {
    process.stdout.write(
      "tsconfig.depcruise.json path aliases are exactly workspace package names.\n",
    );
    return;
  }

  process.stderr.write(`${issues.map((issue) => `- ${issue}`).join("\n")}\n`);
  process.exitCode = 1;
};

if (import.meta.main) main();
