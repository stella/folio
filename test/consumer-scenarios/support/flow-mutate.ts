/**
 * Mutating corpus flows: new flows near ones that reached a new state. A
 * mutant inserts, deletes, swaps or re-seeds steps, drops a step's pinned
 * operations or one of them, or splices in the tail of another corpus flow.
 * Imports nothing from folio.
 */

import {
  type Action,
  type FlowFile,
  type FlowStep,
  flowId,
  LEGACY_ACTIONS,
  TARGETED_ACTIONS,
} from "./flow-file.ts";
import type { Random } from "./random.ts";

export const MUTATIONS = [
  "insert",
  "delete",
  "swap",
  "reseed",
  "unpin",
  "dropOperation",
  "splice",
] as const;
export type Mutation = (typeof MUTATIONS)[number];

const newSeed = (random: Random): number => Math.floor(random.next() * 2 ** 32) >>> 0;

const actionsOf = (flow: FlowFile): readonly Action[] =>
  flow.generation === "legacy" ? LEGACY_ACTIONS : TARGETED_ACTIONS;

/** One mutation of `steps`, or null when it does not apply to them. */
const mutateOnce = (
  flow: FlowFile,
  steps: FlowStep[],
  mutation: Mutation,
  random: Random,
  donors: readonly FlowFile[],
): FlowStep[] | null => {
  const at = random.int(steps.length);
  const step = steps[at] as FlowStep;
  switch (mutation) {
    case "insert": {
      const position = random.int(steps.length + 1);
      const inserted: FlowStep = { action: random.pick(actionsOf(flow)), seed: newSeed(random) };
      return [...steps.slice(0, position), inserted, ...steps.slice(position)];
    }
    case "delete":
      return steps.length < 2 ? null : steps.filter((_, index) => index !== at);
    case "swap": {
      if (steps.length < 2) return null;
      const other = random.int(steps.length);
      if (other === at) return null;
      const swapped = [...steps];
      swapped[at] = steps[other] as FlowStep;
      swapped[other] = step;
      return swapped;
    }
    case "reseed": {
      const reseeded = [...steps];
      reseeded[at] = { action: step.action, seed: newSeed(random) };
      return reseeded;
    }
    case "unpin": {
      const pinned = steps.flatMap((candidate, index) =>
        candidate.operations === undefined ? [] : [index],
      );
      if (pinned.length === 0) return null;
      const index = random.pick(pinned);
      const unpinned = [...steps];
      const { action, seed } = steps[index] as FlowStep;
      unpinned[index] = { action, seed };
      return unpinned;
    }
    case "dropOperation": {
      const pinned = steps.flatMap((candidate, index) =>
        (candidate.operations?.length ?? 0) > 1 ? [index] : [],
      );
      if (pinned.length === 0) return null;
      const index = random.pick(pinned);
      const target = steps[index] as FlowStep;
      const operations = [...(target.operations ?? [])];
      operations.splice(random.int(operations.length), 1);
      const dropped = [...steps];
      dropped[index] = { ...target, operations };
      return dropped;
    }
    case "splice": {
      const donor = donors.length === 0 ? null : random.pick(donors);
      if (donor === null || donor.steps.length === 0) return null;
      const from = random.int(donor.steps.length);
      // A donor on another fixture pins ids this one does not have: keep its steps, redraw.
      const sameDocument =
        donor.fixture === flow.fixture &&
        donor.kind === flow.kind &&
        JSON.stringify(donor.swarm) === JSON.stringify(flow.swarm);
      const tail = donor.steps
        .slice(from)
        .filter((candidate) => actionsOf(flow).includes(candidate.action))
        .map((candidate) =>
          sameDocument ? candidate : { action: candidate.action, seed: candidate.seed },
        );
      return [...steps.slice(0, at + 1), ...tail];
    }
  }
};

/**
 * A mutant of `flow`: one to three mutations, at most `maxSteps` steps,
 * never empty. `donors` are the flows a splice may take a tail from.
 */
export const mutateFlow = (
  flow: FlowFile,
  random: Random,
  { donors = [], maxSteps = 24 }: { donors?: readonly FlowFile[]; maxSteps?: number } = {},
): { flow: FlowFile; mutations: Mutation[] } => {
  let steps = [...flow.steps];
  const applied: Mutation[] = [];
  const count = 1 + random.int(3);
  for (let attempt = 0; applied.length < count && attempt < count * 4; attempt += 1) {
    const mutation = random.pick(MUTATIONS);
    const next = mutateOnce(flow, steps, mutation, random, donors);
    if (next === null || next.length === 0) continue;
    steps = next.slice(0, maxSteps);
    applied.push(mutation);
  }
  if (applied.length === 0) {
    steps = [...steps, { action: random.pick(actionsOf(flow)), seed: newSeed(random) }].slice(
      -maxSteps,
    );
    applied.push("insert");
  }
  return {
    flow: {
      version: 1,
      kind: flow.kind,
      generation: flow.generation,
      fixture: flow.fixture,
      mode: flow.mode,
      seed: flow.seed,
      steps,
      ...(flow.swarm === undefined ? {} : { swarm: flow.swarm }),
      origin: `${applied.join("+")} of ${flowId(flow)}`,
    },
    mutations: applied,
  };
};
