/**
 * Shrinking a flow file to a minimal one that still does what matters:
 * delta debugging over its steps, then over each batch step's operations,
 * then over the steps once more (a step with pinned operations often goes
 * where its drawn one could not). What matters is the caller's predicate:
 * failing with the same fingerprint, or reaching the same new states.
 * Imports nothing from folio.
 */

import { BATCH_ACTIONS, type FlowFile, type FlowStep } from "./flow-file.ts";

/** How much replaying a shrink may spend. */
export type ShrinkBudget = { maxAttempts: number; deadline: number };

export type ShrinkResult = { flow: FlowFile; attempts: number; from: number };

type Tracker = { attempts: number; budget: ShrinkBudget };

const spent = ({ attempts, budget }: Tracker): boolean =>
  attempts >= budget.maxAttempts || Date.now() >= budget.deadline;

/**
 * The shortest subsequence of `items` (at least `min` long) `keeps` still
 * accepts that removing chunks finds: chunks of half the length first, down
 * to single items, until no single item can go.
 */
export const shrinkSequence = async <T>(
  items: readonly T[],
  keeps: (candidate: T[]) => Promise<boolean>,
  tracker: Tracker,
  min = 1,
): Promise<T[]> => {
  let current = [...items];
  let chunk = Math.max(1, Math.floor(current.length / 2));
  while (!spent(tracker)) {
    let removed = false;
    for (let start = 0; start < current.length && !spent(tracker);) {
      const candidate = [...current.slice(0, start), ...current.slice(start + chunk)];
      if (candidate.length < min) {
        start += chunk;
        continue;
      }
      tracker.attempts += 1;
      if (await keeps(candidate)) {
        current = candidate;
        removed = true;
      } else {
        start += chunk;
      }
    }
    if (chunk === 1 && !removed) break;
    if (!removed) chunk = Math.max(1, Math.floor(chunk / 2));
    else chunk = Math.max(1, Math.min(chunk, Math.floor(current.length / 2)));
  }
  return current;
};

const withSteps = (flow: FlowFile, steps: FlowStep[]): FlowFile => ({ ...flow, steps });

/** `flow` with step `index`'s pinned operations cut to the fewest that still fail. */
const shrinkOperations = async (
  flow: FlowFile,
  index: number,
  holds: (candidate: FlowFile) => Promise<boolean>,
  tracker: Tracker,
): Promise<FlowFile> => {
  const step = flow.steps[index] as FlowStep;
  if (!BATCH_ACTIONS.has(step.action) || (step.operations?.length ?? 0) < 2) return flow;
  const withOperations = (operations: FlowStep["operations"]): FlowFile => {
    const steps = [...flow.steps];
    steps[index] = { ...step, operations };
    return withSteps(flow, steps);
  };
  const operations = await shrinkSequence(
    step.operations ?? [],
    (candidate) => holds(withOperations(candidate)),
    tracker,
  );
  return withOperations(operations);
};

/**
 * Shrink `flow`, for which `holds`. `materialize` replays a flow and returns it
 * with every batch step's operations pinned to what the step applied, which
 * turns the second phase's parameters into data it can cut.
 */
export const shrinkFlow = async (
  flow: FlowFile,
  {
    holds,
    materialize,
    budget,
  }: {
    holds: (candidate: FlowFile) => Promise<boolean>;
    materialize: (candidate: FlowFile) => Promise<FlowFile>;
    budget: ShrinkBudget;
  },
): Promise<ShrinkResult> => {
  const tracker: Tracker = { attempts: 0, budget };
  const shrinkSteps = async (current: FlowFile): Promise<FlowFile> =>
    withSteps(
      current,
      await shrinkSequence(current.steps, (steps) => holds(withSteps(current, steps)), tracker),
    );

  let current = await shrinkSteps(flow);
  if (!spent(tracker)) {
    const pinned = await materialize(current);
    tracker.attempts += 1;
    if (await holds(pinned)) {
      current = pinned;
      for (let index = 0; index < current.steps.length && !spent(tracker); index += 1) {
        current = await shrinkOperations(current, index, holds, tracker);
      }
      current = await shrinkSteps(current);
    }
  }
  return {
    flow: { ...current, origin: `shrunk from ${flow.origin ?? "a flow"}` },
    attempts: tracker.attempts,
    from: flow.steps.length,
  };
};
