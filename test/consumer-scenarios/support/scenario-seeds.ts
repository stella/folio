import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import type { FlowKind, Generation } from "./fuzz.ts";

export type PinnedScenarioFlow = {
  seed: number;
  steps: number;
  kind: FlowKind;
  generation: Generation;
  skipKnownIssue?: true;
  title: string;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const parseFlow = (value: unknown): PinnedScenarioFlow => {
  if (!isRecord(value)) {
    throw new TypeError(`Invalid pinned consumer scenario flow: ${JSON.stringify(value)}`);
  }
  const seed = value["seed"];
  const steps = value["steps"];
  const kind = value["kind"];
  const generation = value["generation"];
  const skipKnownIssue = value["skipKnownIssue"];
  const title = value["title"];
  if (
    typeof seed !== "number" ||
    !Number.isSafeInteger(seed) ||
    typeof steps !== "number" ||
    !Number.isSafeInteger(steps) ||
    steps <= 0 ||
    (kind !== "random" && kind !== "collisions") ||
    (generation !== "legacy" && generation !== "targeted") ||
    (skipKnownIssue !== undefined && skipKnownIssue !== true) ||
    typeof title !== "string" ||
    title.length === 0
  ) {
    throw new TypeError(`Invalid pinned consumer scenario flow: ${JSON.stringify(value)}`);
  }
  return {
    seed,
    steps,
    kind,
    generation,
    ...(skipKnownIssue === true ? { skipKnownIssue: true } : {}),
    title,
  };
};

const titleSlug = (title: string): string =>
  title
    .toLowerCase()
    .replaceAll("'", "")
    .replaceAll(/[^a-z0-9]+/gu, "-")
    .replaceAll(/^-|-$/gu, "");

export const scenarioSeedFileName = ({ seed, kind, title }: PinnedScenarioFlow): string =>
  `${String(seed)}-${kind}-${titleSlug(title)}.json`;

export const loadScenarioSeedRegistry = (directory: string): PinnedScenarioFlow[] => {
  const fileNames = readdirSync(directory)
    .filter((fileName) => fileName.endsWith(".json"))
    .sort();
  const flows = fileNames.map((fileName) => {
    const flow = parseFlow(JSON.parse(readFileSync(path.join(directory, fileName), "utf8")));
    const expectedFileName = scenarioSeedFileName(flow);
    if (fileName !== expectedFileName) {
      throw new TypeError(
        `Invalid pinned consumer scenario file name: expected ${expectedFileName}, received ${fileName}`,
      );
    }
    return flow;
  });
  const keys = flows.map(({ seed, kind }) => `${String(seed)}:${kind}`);
  if (new Set(keys).size !== keys.length) {
    throw new TypeError("Invalid pinned consumer scenario registry: duplicate seed and kind");
  }
  return flows;
};

export const PINNED_SCENARIO_FLOWS = loadScenarioSeedRegistry(
  path.join(path.dirname(fileURLToPath(import.meta.url)), "../seeds"),
);
