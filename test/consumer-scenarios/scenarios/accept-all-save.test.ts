import { test } from "node:test";

import { runFlow } from "../support/fuzz.ts";

const GENERATED_REGRESSIONS = [
  { issue: 1751, seed: 1_809_930_855, steps: 20, kind: "collisions" as const },
  { issue: 1773, seed: 746_947_904, steps: 16, kind: "collisions" as const },
  { issue: 1775, seed: 746_948_032, steps: 16, kind: "collisions" as const },
  { issue: 1834, seed: 1_101_252_766, steps: 16, kind: "random" as const },
  { issue: 1836, seed: 1_101_252_863, steps: 16, kind: "random" as const },
  { issue: 1861, seed: 740_097_460, steps: 20, kind: "collisions" as const },
  { issue: 1865, seed: 740_097_627, steps: 20, kind: "random" as const },
  { issue: 1874, seed: 740_097_319, steps: 20, kind: "random" as const },
] as const;

for (const regression of GENERATED_REGRESSIONS) {
  test(
    `#${regression.issue} seed ${regression.seed}: accept-all survives save`,
    { timeout: 60_000 },
    async () => {
      await runFlow(regression.seed, regression.steps, regression.kind);
    },
  );
}

const MINIMIZED_REPLAY_SEEDS = [
  { issue: 1912, seed: 1_492_412_849 },
  { issue: 1915, seed: 774_176_527 },
  { issue: 1916, seed: 1_502_256_441 },
  { issue: 1919, seed: 1_502_256_209 },
  { issue: 1920, seed: 1_026_399_033 },
  { issue: 1928, seed: 1_026_397_717 },
  { issue: 1933, seed: 1_502_256_502 },
] as const;

for (const { issue, seed } of MINIMIZED_REPLAY_SEEDS) {
  test(`#${issue} minimized seed ${seed}: accept-all survives save`, async () => {
    await runFlow(seed, 16, "random");
  });
}
