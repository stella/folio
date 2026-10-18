/** Replays for direct-versus-accepted-tracked equivalence failures. */

import { describe, test } from "node:test";

import { runFlow, type FlowKind } from "../support/fuzz.ts";

type Replay = {
  issue: number;
  seed: number;
  steps: number;
  kind: FlowKind;
  swarm: "enabled" | "disabled";
};

const COMMENT_DELETION_REPLAYS = [
  { issue: 1748, seed: 1809931072, steps: 20, kind: "random", swarm: "enabled" },
  { issue: 1752, seed: 1809930833, steps: 20, kind: "collisions", swarm: "enabled" },
  { issue: 1755, seed: 1809930881, steps: 20, kind: "collisions", swarm: "enabled" },
  { issue: 1763, seed: 746947958, steps: 16, kind: "random", swarm: "disabled" },
  { issue: 1805, seed: 1339716339, steps: 20, kind: "random", swarm: "enabled" },
  { issue: 1812, seed: 1339716521, steps: 20, kind: "random", swarm: "enabled" },
  { issue: 1813, seed: 1339716112, steps: 20, kind: "collisions", swarm: "enabled" },
  { issue: 1816, seed: 1339716517, steps: 20, kind: "random", swarm: "enabled" },
  { issue: 1824, seed: 1101252668, steps: 16, kind: "random", swarm: "disabled" },
  { issue: 1825, seed: 1101252679, steps: 16, kind: "random", swarm: "disabled" },
  { issue: 1840, seed: 1101252700, steps: 16, kind: "collisions", swarm: "disabled" },
  { issue: 1862, seed: 740097170, steps: 20, kind: "collisions", swarm: "enabled" },
  { issue: 1866, seed: 740097276, steps: 20, kind: "collisions", swarm: "enabled" },
  { issue: 1867, seed: 740097637, steps: 20, kind: "random", swarm: "enabled" },
  { issue: 1878, seed: 740097416, steps: 20, kind: "collisions", swarm: "enabled" },
  { issue: 1879, seed: 740097639, steps: 20, kind: "random", swarm: "enabled" },
  { issue: 1892, seed: 14491467, steps: 16, kind: "random", swarm: "disabled" },
] as const satisfies readonly Replay[];

describe("accepted tracked deletions match direct comment lifecycle", () => {
  for (const replay of COMMENT_DELETION_REPLAYS) {
    test(
      `#${replay.issue}: generated ${replay.kind} operation mix (seed ${replay.seed})`,
      { timeout: 30_000 },
      async () => {
        await runFlow(replay.seed, replay.steps, replay.kind, {
          generation: "targeted",
          swarm: replay.swarm,
        });
      },
    );
  }
});
