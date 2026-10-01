/**
 * Flow files: a fuzz flow as data, so it can be replayed without its seed,
 * shrunk step by step, stored in a corpus and mutated.
 *
 * A seeded flow (support/fuzz.ts) draws its fixture, its mode and each step
 * from one generator. A flow file records, per step, the action and the
 * generator's position after the action was drawn (`seed`); replaying the
 * step resumes the generator there, so the step draws what it drew in the
 * seeded run as long as the document is the same. Removing an earlier step
 * changes the document, and the step then draws against what is there. A
 * step may also carry the exact `operations` it applies (a shrunk flow
 * does): the step still draws, to keep its later draws in place, and applies
 * these instead. Imports nothing from folio, so scripts can read it too.
 */

import { createHash } from "node:crypto";

export const FLOW_KINDS = ["random", "collisions"] as const;
export type FlowKind = (typeof FLOW_KINDS)[number];

export const GENERATIONS = ["targeted", "legacy"] as const;
/** How a flow draws its steps; see support/fuzz.ts. */
export type Generation = (typeof GENERATIONS)[number];

/** The steps a legacy flow draws from, in draw order (the order is part of every seed). */
export const LEGACY_ACTIONS = [
  "suggest_changes",
  "suggest_changes",
  "core batch",
  "mistake",
  "add_comment",
  "reply and resolve",
  "accept one",
  "reject one",
  "accept all",
  "reject all",
  "save and reopen",
] as const;

/**
 * A targeted flow's steps: the legacy ones, edits aimed at a header, footer
 * or note, and two more session changes (one step in five changes session).
 */
export const TARGETED_ACTIONS = [
  ...LEGACY_ACTIONS,
  "story batch",
  "story batch",
  "new reviewer",
  "selective save",
] as const;

export type Action = (typeof TARGETED_ACTIONS)[number];

const ACTIONS: ReadonlySet<string> = new Set(TARGETED_ACTIONS);

/** The steps that draw a batch of operations, which a step may pin. */
export const BATCH_ACTIONS: ReadonlySet<Action> = new Set([
  "suggest_changes",
  "core batch",
  "story batch",
]);

export type FlowStep = {
  action: Action;
  /** The generator's position when the step starts drawing. */
  seed: number;
  /** The operations a batch step applies instead of the ones it draws. */
  operations?: Record<string, unknown>[];
};

export type FlowFile = {
  version: 1;
  kind: FlowKind;
  generation: Generation;
  fixture: string;
  mode: string;
  /** The seed the flow's metamorphic relations sample with (a seeded flow's own seed). */
  seed: number;
  steps: FlowStep[];
  /** Generated operation kinds enabled for this swarm; absent means all kinds. */
  swarm?: string[];
  /** What the flow is: the seed it was drawn from, a corpus mutation, a shrink. */
  origin?: string;
  /** For a checked-in flow: what it guards. */
  title?: string;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isUint32 = (value: unknown): value is number =>
  typeof value === "number" && Number.isInteger(value) && value >= 0 && value < 2 ** 32;

const parseStep = (value: unknown, index: number): FlowStep => {
  if (!isRecord(value) || typeof value["action"] !== "string" || !ACTIONS.has(value["action"])) {
    throw new TypeError(`flow step ${index}: unknown action ${JSON.stringify(value)}`);
  }
  const seed = value["seed"];
  if (!isUint32(seed)) {
    throw new TypeError(`flow step ${index}: seed must be an unsigned 32-bit integer`);
  }
  const action = value["action"] as Action;
  const operations = value["operations"];
  if (operations === undefined) return { action, seed };
  if (
    !BATCH_ACTIONS.has(action) ||
    !Array.isArray(operations) ||
    !operations.every((operation) => isRecord(operation) && typeof operation["type"] === "string")
  ) {
    throw new TypeError(`flow step ${index}: operations need a batch step and a type each`);
  }
  return { action, seed, operations: operations as Record<string, unknown>[] };
};

/** Read a flow file; throws a TypeError naming what is wrong. */
export const parseFlowFile = (value: unknown): FlowFile => {
  if (!isRecord(value) || value["version"] !== 1) {
    throw new TypeError("flow file: expected version 1");
  }
  const { kind, generation, fixture, mode, seed, steps, origin, title, swarm } = value;
  if (!(FLOW_KINDS as readonly unknown[]).includes(kind)) {
    throw new TypeError(`flow file: unknown kind ${JSON.stringify(kind)}`);
  }
  if (!(GENERATIONS as readonly unknown[]).includes(generation)) {
    throw new TypeError(`flow file: unknown generation ${JSON.stringify(generation)}`);
  }
  if (typeof fixture !== "string" || typeof mode !== "string") {
    throw new TypeError("flow file: fixture and mode must be strings");
  }
  if (typeof seed !== "number" || !Number.isSafeInteger(seed)) {
    throw new TypeError("flow file: seed must be an integer");
  }
  if (!Array.isArray(steps) || steps.length === 0) {
    throw new TypeError("flow file: expected at least one step");
  }
  if (
    swarm !== undefined &&
    (!Array.isArray(swarm) ||
      swarm.length === 0 ||
      !swarm.every((type): type is string => typeof type === "string" && type.length > 0) ||
      new Set(swarm).size !== swarm.length)
  ) {
    throw new TypeError("flow file: swarm must be a nonempty list of distinct operation kinds");
  }
  return {
    version: 1,
    kind: kind as FlowKind,
    generation: generation as Generation,
    fixture,
    mode,
    seed,
    steps: steps.map(parseStep),
    ...(swarm === undefined ? {} : { swarm }),
    ...(typeof origin === "string" ? { origin } : {}),
    ...(typeof title === "string" ? { title } : {}),
  };
};

/** Short, stable id of what a flow does (not where it came from). */
export const flowId = (flow: FlowFile): string =>
  createHash("sha256")
    .update(
      JSON.stringify([
        flow.kind,
        flow.generation,
        flow.fixture,
        flow.mode,
        flow.seed,
        flow.steps,
        ...(flow.swarm === undefined ? [] : [flow.swarm]),
      ]),
    )
    .digest("hex")
    .slice(0, 16);

/**
 * The operations a flow applies, step by step: a batch step's operation
 * types, any other step's action; ids, text and positions left out. Part of
 * a shrunk failure's fingerprint (support/failure-fingerprints.ts).
 */
export const flowShape = (flow: FlowFile): string =>
  flow.steps
    .map((step) =>
      step.operations === undefined || step.operations.length === 0
        ? step.action
        : [...new Set(step.operations.map((operation) => String(operation["type"])))]
            .sort()
            .join("+"),
    )
    .join(" > ");

/** A single-line JSON of `flow`, for an environment variable. */
export const compactFlow = (flow: FlowFile): string => JSON.stringify(flow);
