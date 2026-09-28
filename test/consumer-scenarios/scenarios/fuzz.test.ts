/**
 * Pinned regression flows from scenario-seeds.json run before the generated
 * random and collision runs. The generated seed and run counts are controlled
 * by FOLIO_SCENARIO_SEED, FOLIO_SCENARIO_FUZZ_RUNS,
 * FOLIO_SCENARIO_COLLISION_RUNS and FOLIO_SCENARIO_FUZZ_STEPS.
 */

import { after, test } from "node:test";

import { reportScenarioFailure, shellQuote } from "../support/failure-fingerprints.ts";
import { describeFlow, type FlowKind, runFlow } from "../support/fuzz.ts";
import { ENABLED_RELATIONS, relationSummary } from "../support/metamorphic.ts";
import { KNOWN_FAILING_FLOWS } from "../support/known-issues.ts";
import { PINNED_SCENARIO_FLOWS } from "../support/scenario-seeds.ts";

const integer = (value: string | undefined, fallback: number): number => {
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed >= 0 ? parsed : fallback;
};

const SEED = integer(process.env["FOLIO_SCENARIO_SEED"], 20_260_926);
const RUNS = integer(process.env["FOLIO_SCENARIO_FUZZ_RUNS"], 12);
const STEPS = integer(process.env["FOLIO_SCENARIO_FUZZ_STEPS"], 10);
const COLLISION_RUNS = integer(process.env["FOLIO_SCENARIO_COLLISION_RUNS"], 8);

const relationRepro = ["FOLIO_SCENARIO_RELATIONS", "FOLIO_SCENARIO_RELATIONS_DEPTH"]
  .flatMap((name) => {
    const value = process.env[name];
    return value === undefined ? [] : [`${name}=${shellQuote(value)}`];
  })
  .join(" ");

const generatedFlowTest = (kind: FlowKind, run: number, seed: number) => {
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
      ? `FOLIO_SCENARIO_FUZZ_RUNS=${run + 1} FOLIO_SCENARIO_COLLISION_RUNS=0`
      : `FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=${run + 1}`;
  const label = kind === "random" ? "fuzz" : "collision";
  const repro = `FOLIO_SCENARIO_SEED=${SEED} FOLIO_SCENARIO_FUZZ_STEPS=${STEPS} ${runs} ${relationRepro} bun scripts/consumer-scenarios.ts --only '^${label} run ${run} \\('`;
  test(
    `${label} run ${run} (seed ${seed}): ${fixture} / ${mode}, ${STEPS} steps`,
    known ? { skip: `reproduces ${known.finding}; runs in known-issues.test.ts` } : {},
    async () => {
      try {
        await runFlow(seed, STEPS, kind);
      } catch (error) {
        reportScenarioFailure({
          test: `consumer flow ${fixture} / ${mode}`,
          seed,
          repro,
          failure: error,
        });
      }
    },
  );
};

after(() => {
  console.log(relationSummary());
});

// Keep pinned cases ahead of generated tests so every run checks known seeds first.
for (const flow of PINNED_SCENARIO_FLOWS) {
  const known = flow.skipKnownIssue
    ? KNOWN_FAILING_FLOWS.find(
        (knownFlow) =>
          knownFlow.seed === flow.seed &&
          knownFlow.steps === flow.steps &&
          (knownFlow.kind ?? "random") === flow.kind &&
          (knownFlow.generation ?? "targeted") === flow.generation &&
          (knownFlow.relation === undefined || ENABLED_RELATIONS.has(knownFlow.relation)),
      )
    : undefined;
  test(
    `${flow.kind} flow with seed ${flow.seed} (${flow.steps} steps) ${flow.title}`,
    known ? { skip: `reproduces ${known.finding}; runs in known-issues.test.ts` } : {},
    async () => {
      try {
        await runFlow(flow.seed, flow.steps, flow.kind, { generation: flow.generation });
      } catch (error) {
        const { fixture, mode } = describeFlow(flow.seed, flow.kind);
        const repro = `FOLIO_SCENARIO_SEED=${flow.seed} FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=0 ${relationRepro} bun scripts/consumer-scenarios.ts --only '^${flow.kind} flow with seed ${flow.seed} \\('`;
        reportScenarioFailure({
          test: `consumer flow ${fixture} / ${mode}`,
          seed: flow.seed,
          repro,
          failure: error,
        });
      }
    },
  );
}

for (let run = 0; run < RUNS; run += 1) generatedFlowTest("random", run, SEED + run);
for (let run = 0; run < COLLISION_RUNS; run += 1) generatedFlowTest("collisions", run, SEED + run);
