/** Seeded operation subsets, drawn independently of the flow's random stream. */
import { createRandom } from "./random.ts";

/** Collision batches must remain intact; dropping a member changes the collision. */
export const swarmIncludesBatch = (
  swarm: readonly string[] | undefined,
  operations: readonly { type: string }[],
): boolean => swarm === undefined || operations.every(({ type }) => swarm.includes(type));

export const drawSwarm = (seed: number, kinds: readonly string[]): string[] => {
  if (kinds.length === 0 || new Set(kinds).size !== kinds.length) {
    throw new TypeError("swarm requires distinct operation kinds and at least one kind");
  }
  const random = createRandom(seed ^ 0x53574152);
  const enabled = kinds.filter(() => random.chance(0.5));
  if (enabled.length === 0) enabled.push(random.pick(kinds));
  if (enabled.length === kinds.length && kinds.length > 1) {
    enabled.splice(random.int(enabled.length), 1);
  }
  return enabled;
};
