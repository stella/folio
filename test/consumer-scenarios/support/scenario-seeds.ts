import registry from "../scenario-seeds.json" with { type: "json" };

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

export const parseScenarioSeedRegistry = (value: unknown): PinnedScenarioFlow[] => {
  if (!isRecord(value) || !Array.isArray(value["flows"])) {
    throw new TypeError("Invalid pinned consumer scenario registry: expected a flows array");
  }
  const flows = value["flows"].map(parseFlow);
  const keys = flows.map(({ seed, steps, kind, generation }) =>
    [seed, steps, kind, generation].join(":"),
  );
  if (new Set(keys).size !== keys.length) {
    throw new TypeError("Invalid pinned consumer scenario registry: duplicate flow");
  }
  return flows;
};

export const PINNED_SCENARIO_FLOWS = parseScenarioSeedRegistry(registry);
