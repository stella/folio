/**
 * An optional external corpus, read from a directory the runner is pointed at.
 *
 * Nothing here is committed. Third-party comparison corpora carry their own
 * licences and their own idea of what a "revision" is, so the repository keeps
 * its own small authored fixtures and treats an imported corpus purely as a
 * measurement: fetch it locally, point `--corpus` at it, read the numbers.
 * Without `--corpus` these suites do not run and the rest of the benchmark is
 * unaffected.
 *
 * The contract is one file. `pairs.json` in the corpus directory:
 *
 * ```json
 * [{ "id": "case-1", "base": "a.docx", "target": "b.docx", "expectedChanges": 4 }]
 * ```
 *
 * `expectedChanges` is optional and is never asserted on: a revision count is
 * relative to the engine that produced it, so it is read as a change detector
 * and every divergence is reported for a human to explain. Failing a
 * `pairs.json`, a directory of `<case>/baseline.docx` plus `<case>/candidate.docx`
 * is discovered automatically.
 */

import { existsSync, readFileSync, readdirSync } from "node:fs";
import path from "node:path";

export type CorpusPair = {
  id: string;
  basePath: string;
  targetPath: string;
  /** The source corpus's own revision count, when it publishes one. */
  expectedChanges: number | null;
};

const PAIRS_FILE = "pairs.json";

type PairsEntry = {
  id?: unknown;
  base?: unknown;
  target?: unknown;
  expectedChanges?: unknown;
};

const readPairsFile = (directory: string): CorpusPair[] => {
  const entries: unknown = JSON.parse(readFileSync(path.join(directory, PAIRS_FILE), "utf8"));
  if (!Array.isArray(entries)) {
    throw new Error(`${PAIRS_FILE} must hold an array of pairs.`);
  }
  return entries.map((entry: PairsEntry, index) => {
    const { id, base, target, expectedChanges } = entry;
    if (typeof base !== "string" || typeof target !== "string") {
      throw new Error(`${PAIRS_FILE} entry ${String(index)} needs string base and target paths.`);
    }
    return {
      id: typeof id === "string" ? id : `${base} vs ${target}`,
      basePath: path.resolve(directory, base),
      targetPath: path.resolve(directory, target),
      expectedChanges: typeof expectedChanges === "number" ? expectedChanges : null,
    };
  });
};

const BASE_NAMES = Object.freeze(["baseline", "before", "base"] as const);
const TARGET_NAMES = Object.freeze(["candidate", "after", "target"] as const);
const EXTENSIONS = Object.freeze([".docx", ".docm"] as const);

const findNamed = (
  directory: string,
  names: readonly string[],
  files: readonly string[],
): string | null => {
  for (const name of names) {
    for (const extension of EXTENSIONS) {
      if (files.includes(`${name}${extension}`)) {
        return path.join(directory, `${name}${extension}`);
      }
    }
  }
  return null;
};

const discoverPairedDirectories = (root: string): CorpusPair[] => {
  const pairs: CorpusPair[] = [];
  for (const entry of readdirSync(root, { withFileTypes: true }).toSorted((left, right) =>
    left.name.localeCompare(right.name),
  )) {
    if (!entry.isDirectory()) {
      continue;
    }
    const directory = path.join(root, entry.name);
    const files = readdirSync(directory);
    const basePath = findNamed(directory, BASE_NAMES, files);
    const targetPath = findNamed(directory, TARGET_NAMES, files);
    if (basePath !== null && targetPath !== null) {
      pairs.push({ id: entry.name, basePath, targetPath, expectedChanges: null });
    }
  }
  return pairs;
};

export type LoadCorpusResult =
  | { status: "loaded"; pairs: readonly CorpusPair[] }
  | { status: "empty"; detail: string };

export const loadCorpus = (directory: string): LoadCorpusResult => {
  if (!existsSync(directory)) {
    return { status: "empty", detail: `no such directory: ${directory}` };
  }
  const pairs = existsSync(path.join(directory, PAIRS_FILE))
    ? readPairsFile(directory)
    : discoverPairedDirectories(directory);
  const usable = pairs.filter(
    ({ basePath, targetPath }) => existsSync(basePath) && existsSync(targetPath),
  );
  return usable.length === 0
    ? { status: "empty", detail: `no readable pairs under ${directory}` }
    : { status: "loaded", pairs: usable };
};

export const readDocument = (file: string): ArrayBuffer => {
  const bytes = readFileSync(file);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};
