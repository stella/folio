/**
 * Seeded random flows (support/fuzz.ts) over the synthetic documents.
 *
 * FOLIO_SCENARIO_SEED fixes the flows (the runner prints it; run `i` uses
 * seed + i), FOLIO_SCENARIO_FUZZ_RUNS / FOLIO_SCENARIO_FUZZ_STEPS size the
 * search, and FOLIO_SCENARIO_COLLISION_RUNS the collision flows (support/fuzz.ts)
 * that follow them, seeded from the same base. Both use the targeted generation
 * (support/fuzz.ts); seeds pinned before it existed replay on the legacy one.
 * A flow that reproduces a known finding runs in known-issues.test.ts
 * instead, as an expected failure.
 * FOLIO_SCENARIO_RELATIONS / FOLIO_SCENARIO_RELATIONS_DEPTH pick the
 * metamorphic relations checked after every step (support/metamorphic.ts);
 * the run ends by printing which ran.
 */

import { after, test } from "node:test";

import { describeFlow, type FlowKind, runFlow } from "../support/fuzz.ts";
import { KNOWN_FAILING_FLOWS } from "../support/known-issues.ts";
import { ENABLED_RELATIONS, relationSummary } from "../support/metamorphic.ts";

const integer = (value: string | undefined, fallback: number): number => {
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : fallback;
};

const SEED = integer(process.env["FOLIO_SCENARIO_SEED"], 20_260_926);
const RUNS = integer(process.env["FOLIO_SCENARIO_FUZZ_RUNS"], 12);
const STEPS = integer(process.env["FOLIO_SCENARIO_FUZZ_STEPS"], 10);
const COLLISION_RUNS = integer(process.env["FOLIO_SCENARIO_COLLISION_RUNS"], 8);

const flowTest = (kind: FlowKind, run: number, seed: number) => {
  const { fixture, mode } = describeFlow(seed, kind);
  const known = KNOWN_FAILING_FLOWS.find(
    (flow) =>
      flow.seed === seed &&
      flow.steps === STEPS &&
      (flow.kind ?? "random") === kind &&
      (flow.generation ?? "targeted") === "targeted" &&
      (flow.relation === undefined || ENABLED_RELATIONS.has(flow.relation)),
  );
  const runs =
    kind === "random"
      ? `FOLIO_SCENARIO_FUZZ_RUNS=${run + 1}`
      : `FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=${run + 1}`;
  test(
    `${kind === "random" ? "fuzz" : "collision"} run ${run} (seed ${seed}): ${fixture} / ${mode}, ${STEPS} steps`,
    known ? { skip: `reproduces ${known.finding}; runs in known-issues.test.ts` } : {},
    async () => {
      try {
        await runFlow(seed, STEPS, kind);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        throw new Error(
          `Replay: FOLIO_SCENARIO_SEED=${SEED} ${runs} bun run test:consumer-scenarios -- fuzz.test.ts\n${message}`,
          { cause: error },
        );
      }
    },
  );
};

after(() => {
  console.log(relationSummary());
});

for (let run = 0; run < RUNS; run += 1) flowTest("random", run, SEED + run);
for (let run = 0; run < COLLISION_RUNS; run += 1) flowTest("collisions", run, SEED + run);

// Flows that once lost accepted suggestions on save (a suggested paragraph
// deleted again, then every suggestion accepted); kept as fixed seeds on the
// generation that found them.
for (const [seed, steps] of [
  [20_260_933, 10],
  [99, 15],
  [101, 15],
] as const) {
  test(`suggested flow with seed ${seed} (${steps} steps) saves what the reviewer shows`, () =>
    runFlow(seed, steps, "random", { generation: "legacy" }));
}

// Fixed reader-stability failures after table insertion or a paragraph merge:
// live preview formatting must match the reopened package.
for (const { seed, kind, steps } of [
  { seed: 20_260_935, kind: "random", steps: 10 },
  { seed: 20_260_932, kind: "collisions", steps: 10 },
  { seed: 20_260_931, kind: "random", steps: 4 },
] as const) {
  test(`fixed ${kind} flow with seed ${seed} keeps live block fields across save`, () =>
    runFlow(seed, steps, kind, { generation: "targeted" }));
}

// Collision flows that once rewrote a paragraph pending deletion into the one
// after it (1088), and read a deleted pending insertion back as a deletion
// alone after a save (1185); kept as fixed seeds on the generation that
// found them.
for (const seed of [1088, 1185]) {
  const known = KNOWN_FAILING_FLOWS.find(
    (flow) =>
      flow.seed === seed &&
      flow.steps === 10 &&
      flow.kind === "collisions" &&
      flow.generation === "legacy" &&
      (flow.relation === undefined || ENABLED_RELATIONS.has(flow.relation)),
  );
  test(
    `collision flow with seed ${seed} (10 steps) does what it asked and saves it`,
    known ? { skip: `reproduces ${known.finding}; runs in known-issues.test.ts` } : {},
    () => runFlow(seed, 10, "collisions", { generation: "legacy" }),
  );
}

test("collision flow with seed 7310028 keeps comment ranges stable on save", () =>
  runFlow(7_310_028, 10, "collisions", { generation: "legacy" }));

// Deleting a block does not remove a suggested comment or a comment whose
// anchor continues into a surviving block.
for (const { seed, kind } of [
  { seed: 20_260_933, kind: "collisions" },
  { seed: 20_260_937, kind: "random" },
] as const) {
  const known = KNOWN_FAILING_FLOWS.find(
    (flow) =>
      flow.seed === seed &&
      flow.steps === 10 &&
      (flow.kind ?? "random") === kind &&
      (flow.generation ?? "targeted") === "targeted" &&
      (flow.relation === undefined || ENABLED_RELATIONS.has(flow.relation)),
  );
  test(
    `${kind} flow with seed ${seed} keeps surviving comments`,
    known ? { skip: `reproduces ${known.finding}; runs in known-issues.test.ts` } : {},
    () => runFlow(seed, 10, kind),
  );
}
