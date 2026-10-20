import { test } from "node:test";

import { runFlow } from "../support/fuzz.ts";

/** Generated operation mixes whose save/reopen comparison exposed non-canonical empty carriers. */
const GENERATED_REGRESSIONS = [
  { issue: 1808, seed: 1_339_715_993, steps: 20, kind: "random" as const },
  { issue: 1922, seed: 1_502_256_685, steps: 20, kind: "random" as const },
  { issue: 1927, seed: 1_502_256_226, steps: 20, kind: "random" as const },
  { issue: 1927, seed: 1_502_256_688, steps: 20, kind: "random" as const },
  { issue: 1934, seed: 1_502_256_418, steps: 20, kind: "random" as const },
  { issue: 1936, seed: 1_502_256_382, steps: 20, kind: "random" as const },
] as const;

for (const regression of GENERATED_REGRESSIONS) {
  test(
    `#${regression.issue} seed ${regression.seed}: reject-all survives save`,
    { timeout: 60_000 },
    async () => {
      await runFlow(regression.seed, regression.steps, regression.kind);
    },
  );
}
