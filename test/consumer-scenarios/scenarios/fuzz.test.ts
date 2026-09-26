/**
 * Seeded random flows (support/fuzz.ts) over the synthetic documents.
 *
 * FOLIO_SCENARIO_SEED fixes the flows (the runner prints it; run `i` uses
 * seed + i), FOLIO_SCENARIO_FUZZ_RUNS / FOLIO_SCENARIO_FUZZ_STEPS size the
 * search. A flow that reproduces a known finding runs in known-issues.test.ts
 * instead, as an expected failure.
 */

import { test } from "node:test";

import { describeFlow, runFlow } from "../support/fuzz.ts";
import { KNOWN_FAILING_FLOWS } from "../support/known-issues.ts";

const integer = (value: string | undefined, fallback: number): number => {
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : fallback;
};

const SEED = integer(process.env["FOLIO_SCENARIO_SEED"], 20_260_926);
const RUNS = integer(process.env["FOLIO_SCENARIO_FUZZ_RUNS"], 12);
const STEPS = integer(process.env["FOLIO_SCENARIO_FUZZ_STEPS"], 10);

for (let run = 0; run < RUNS; run += 1) {
  const seed = SEED + run;
  const { fixture, mode } = describeFlow(seed);
  const known = KNOWN_FAILING_FLOWS.find((flow) => flow.seed === seed && flow.steps === STEPS);
  test(
    `fuzz run ${run} (seed ${seed}): ${fixture} / ${mode}, ${STEPS} steps`,
    known ? { skip: `reproduces ${known.finding}; runs in known-issues.test.ts` } : {},
    async () => {
      try {
        await runFlow(seed, STEPS);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        throw new Error(
          `Replay: FOLIO_SCENARIO_SEED=${SEED} FOLIO_SCENARIO_FUZZ_RUNS=${run + 1} bun run test:consumer-scenarios -- fuzz.test.ts\n${message}`,
          { cause: error },
        );
      }
    },
  );
}
