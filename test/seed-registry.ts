import { Result, TaggedError } from "better-result";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import ts from "typescript";
import type { PinnedSeed } from "./property-testing";

/**
 * One flat JSON file per test file, named encodeURIComponent(repo-relative path)
 * + `.json`. Decoding and re-encoding must match exactly: `%2F` is canonical,
 * alternate escape spellings are refused. Each object maps test titles to seeds.
 */
export const PROPERTY_SEEDS_FILE = "test/property-seeds";
const LEGACY_FILE = "test/property-seeds.json";
const REPO_ROOT = path.resolve(import.meta.dir, "..");

type SeedRegistry = Record<string, readonly PinnedSeed[]>;
type SeedFiles = Record<string, SeedRegistry>;

class SeedRegistryError extends TaggedError("SeedRegistryError")<{ message: string }> {}
const invalid = (message: string): never => {
  throw new SeedRegistryError({ message });
};

const testFilePath = (file: string): string => {
  if (
    !file ||
    file.includes("\\") ||
    path.posix.isAbsolute(file) ||
    path.posix.normalize(file) !== file ||
    file.startsWith("../")
  )
    return invalid(`Invalid seed test path: ${file}`);
  return file;
};

export const seedFileFor = (file: string): string =>
  `${PROPERTY_SEEDS_FILE}/${encodeURIComponent(testFilePath(file))}.json`;

export const testFileForSeedFile = (file: string): string => {
  if (!file.startsWith(`${PROPERTY_SEEDS_FILE}/`) || !file.endsWith(".json"))
    return invalid(`Invalid seed registry path: ${file}`);
  const decoded = Result.try(() =>
    decodeURIComponent(file.slice(PROPERTY_SEEDS_FILE.length + 1, -5)),
  );
  if (decoded.isErr()) return invalid(`Invalid seed path encoding: ${file}`);
  if (seedFileFor(decoded.value) !== file)
    return invalid(`Noncanonical seed path encoding: ${file}`);
  return decoded.value;
};

const compareNames = (left: string, right: string): number => {
  if (left < right) return -1;
  if (left > right) return 1;
  return 0;
};

export const splitPinnedSeeds = (registry: SeedRegistry): SeedFiles => {
  const files = new Map<string, Map<string, readonly PinnedSeed[]>>();
  for (const [key, seeds] of Object.entries(registry)) {
    const separator = key.indexOf("::");
    if (separator < 1 || separator + 2 === key.length) return invalid(`Invalid seed key: ${key}`);
    const file = seedFileFor(key.slice(0, separator));
    const titles = files.get(file) ?? new Map<string, readonly PinnedSeed[]>();
    titles.set(key.slice(separator + 2), seeds);
    files.set(file, titles);
  }
  return Object.fromEntries(
    [...files]
      .toSorted(([a], [b]) => compareNames(a, b))
      .map(
        ([file, titles]) =>
          [
            file,
            Object.fromEntries([...titles].toSorted(([a], [b]) => compareNames(a, b))),
          ] as const,
      ),
  );
};

export const joinPinnedSeeds = (files: SeedFiles): SeedRegistry =>
  Object.fromEntries(
    Object.entries(files).flatMap(([file, titles]) => {
      const testFile = testFileForSeedFile(file);
      return Object.entries(titles).map(
        ([title, seeds]) => [`${testFile}::${title}`, seeds] as const,
      );
    }),
  );

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const isPinnedSeed = (value: unknown): value is PinnedSeed => {
  if (!isRecord(value)) return false;
  const expected = value["expectedFailure"];
  return (
    typeof value["seed"] === "number" &&
    typeof value["note"] === "string" &&
    typeof value["date"] === "string" &&
    (value["path"] === undefined || typeof value["path"] === "string") &&
    (expected === undefined ||
      (isRecord(expected) &&
        typeof expected["family"] === "string" &&
        typeof expected["fingerprint"] === "string"))
  );
};

/** Parse JSON before inspecting its AST so escaped duplicate keys cannot be overwritten. */
export const parseSeedTitles = (source: string, file: string): SeedRegistry => {
  const result = Result.try(() => JSON.parse(source));
  if (result.isErr()) return invalid(`Invalid seed JSON in ${file}`);
  const parsed: unknown = result.value;
  const ast = ts.createSourceFile(file, source, ts.ScriptTarget.Latest, true, ts.ScriptKind.JSON);
  const visit = (node: ts.Node): void => {
    if (ts.isObjectLiteralExpression(node)) {
      const keys = new Set<string>();
      for (const property of node.properties) {
        if (!ts.isPropertyAssignment(property) || !ts.isStringLiteral(property.name))
          return invalid(`Invalid seed object in ${file}`);
        const key = property.name.text;
        if (keys.has(key)) return invalid(`Duplicate seed key ${JSON.stringify(key)} in ${file}`);
        keys.add(key);
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(ast);
  if (!isRecord(parsed)) return invalid(`Seed titles must be an object in ${file}`);
  const entries: [string, readonly PinnedSeed[]][] = [];
  for (const [title, seeds] of Object.entries(parsed)) {
    if (!title || !Array.isArray(seeds) || !seeds.every(isPinnedSeed))
      return invalid(`Invalid seed entries for ${JSON.stringify(title)} in ${file}`);
    entries.push([title, seeds]);
  }
  return Object.fromEntries(entries);
};

export const readSeedRegistry = (repoRoot = REPO_ROOT): SeedRegistry => {
  if (existsSync(path.join(repoRoot, LEGACY_FILE)))
    return invalid(`${LEGACY_FILE} must not exist; use per-test-file seeds`);
  const files: [string, SeedRegistry][] = [];
  for (const entry of readdirSync(path.join(repoRoot, PROPERTY_SEEDS_FILE), {
    withFileTypes: true,
  }).toSorted((a, b) => compareNames(a.name, b.name))) {
    const file = `${PROPERTY_SEEDS_FILE}/${entry.name}`;
    testFileForSeedFile(file);
    if (!entry.isFile()) return invalid(`Seed registry entry must be a regular file: ${file}`);
    files.push([file, parseSeedTitles(readFileSync(path.join(repoRoot, file), "utf8"), file)]);
  }
  return joinPinnedSeeds(Object.fromEntries(files));
};
