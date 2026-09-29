/**
 * The corpus of interesting flows: flow files that reached a state no earlier
 * flow reached, each kept with its signature (what `runFlow` collects with
 * `signature: true`: coverage cells, step outcomes, document structures). A
 * continuous fuzz run loads it, mutates its flows alongside fresh ones, and
 * adds every flow whose signature has an element the corpus lacks. The CI
 * cache carries it from run to run; test/consumer-scenarios/flows/ holds the
 * small, minimized part that is checked in. Imports nothing from folio.
 */

import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { join } from "node:path";

import { type FlowFile, flowId, parseFlowFile } from "./flow-file.ts";

export type CorpusEntry = { id: string; flow: FlowFile; signature: string[]; found: string };

export type Corpus = {
  /** Where new entries are written, or null to keep them in memory. */
  dir: string | null;
  entries: CorpusEntry[];
  /** How many entries carry each signature element. */
  counts: Map<string, number>;
};

const count = (corpus: Corpus, entry: CorpusEntry, delta: 1 | -1): void => {
  for (const element of entry.signature) {
    const next = (corpus.counts.get(element) ?? 0) + delta;
    if (next <= 0) corpus.counts.delete(element);
    else corpus.counts.set(element, next);
  }
};

const readEntry = (file: string): CorpusEntry | null => {
  try {
    const value = JSON.parse(readFileSync(file, "utf8")) as Record<string, unknown>;
    const flow = parseFlowFile(value["flow"]);
    const signature = value["signature"];
    if (!Array.isArray(signature) || !signature.every((item) => typeof item === "string")) {
      return null;
    }
    return {
      id: flowId(flow),
      flow,
      signature,
      found: typeof value["found"] === "string" ? value["found"] : "",
    };
  } catch {
    // A corrupt cache entry is dropped, never fatal.
    return null;
  }
};

/** The corpus in `dir` (created if missing), plus `seeds` with no signature yet. */
export const loadCorpus = (dir: string | null, seeds: readonly FlowFile[] = []): Corpus => {
  const corpus: Corpus = { dir, entries: [], counts: new Map() };
  const seen = new Set<string>();
  const add = (entry: CorpusEntry): void => {
    if (seen.has(entry.id)) return;
    seen.add(entry.id);
    corpus.entries.push(entry);
    count(corpus, entry, 1);
  };
  if (dir !== null) {
    mkdirSync(dir, { recursive: true });
    for (const file of readdirSync(dir)
      .filter((name) => name.endsWith(".json"))
      .sort()) {
      const entry = readEntry(join(dir, file));
      if (entry !== null) add(entry);
    }
  }
  for (const flow of seeds) add({ id: flowId(flow), flow, signature: [], found: "" });
  return corpus;
};

/** The elements of `signature` no corpus entry has. */
export const newElements = (corpus: Corpus, signature: readonly string[]): string[] =>
  [...new Set(signature)].filter((element) => !corpus.counts.has(element));

/**
 * Add `flow` when its signature has an element the corpus lacks; returns
 * those elements (empty: not added).
 */
export const admit = (
  corpus: Corpus,
  flow: FlowFile,
  signature: readonly string[],
  found = new Date().toISOString(),
): string[] => {
  const fresh = newElements(corpus, signature);
  const id = flowId(flow);
  if (fresh.length === 0 || corpus.entries.some((entry) => entry.id === id)) return [];
  const entry: CorpusEntry = { id, flow, signature: [...new Set(signature)].sort(), found };
  corpus.entries.push(entry);
  count(corpus, entry, 1);
  if (corpus.dir !== null) {
    writeFileSync(
      join(corpus.dir, `${id}.json`),
      `${JSON.stringify({ flow, signature: entry.signature, found })}\n`,
    );
  }
  return fresh;
};

/**
 * Keep at most `max` entries: drop the oldest whose every element another
 * entry also has, then (still over) the oldest. Returns how many went.
 */
export const prune = (corpus: Corpus, max: number): number => {
  let removed = 0;
  const drop = (entry: CorpusEntry): void => {
    corpus.entries.splice(corpus.entries.indexOf(entry), 1);
    count(corpus, entry, -1);
    if (corpus.dir !== null) {
      const file = join(corpus.dir, `${entry.id}.json`);
      if (existsSync(file)) rmSync(file);
    }
    removed += 1;
  };
  const byAge = () =>
    corpus.entries
      .filter((entry) => entry.found !== "")
      .sort((a, b) => a.found.localeCompare(b.found));
  for (const entry of byAge()) {
    if (corpus.entries.length <= max) break;
    if (entry.signature.every((element) => (corpus.counts.get(element) ?? 0) > 1)) drop(entry);
  }
  for (const entry of byAge()) {
    if (corpus.entries.length <= max) break;
    drop(entry);
  }
  return removed;
};

/** The checked-in flows in `dir` (test/consumer-scenarios/flows), sorted by file name. */
export const readFlowDir = (dir: string): { file: string; flow: FlowFile }[] =>
  existsSync(dir)
    ? readdirSync(dir)
        .filter((name) => name.endsWith(".json"))
        .sort()
        .map((file) => ({
          file,
          flow: parseFlowFile(JSON.parse(readFileSync(join(dir, file), "utf8"))),
        }))
    : [];
