import { test } from "node:test";

import { runFlow } from "../support/fuzz.ts";

const PINNED_CARRIER_FLOWS = [
  { seed: 740_097_579, steps: 20 },
  { seed: 746_947_976, steps: 16 },
  { seed: 1_809_931_060, steps: 20 },
] as const;

for (const { seed, steps } of PINNED_CARRIER_FLOWS) {
  test(
    `direct and accepted tracked paragraph carriers agree for seed ${seed}`,
    { timeout: 120_000 },
    async () => {
      await runFlow(seed, steps, "random");
    },
  );
}

test(
  "generated table and heading edits keep direct and accepted tracked carriers equivalent",
  { timeout: 120_000 },
  async () => {
    // These targeted generated flows exercise deleteBlock, splitBlock,
    // mergeBlockWithNext and insertTable against the table/heading fixtures.
    for (let seed = 31_337; seed < 31_347; seed += 1) {
      await runFlow(seed, 24, "random");
    }
  },
);
