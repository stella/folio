/**
 * Pinned regression flows from scenario-seeds.json run before the generated
 * random and collision runs. The generated seed and run counts are controlled
 * by FOLIO_SCENARIO_SEED, FOLIO_SCENARIO_FUZZ_RUNS,
 * FOLIO_SCENARIO_COLLISION_RUNS and FOLIO_SCENARIO_FUZZ_STEPS.
 * FOLIO_SCENARIO_FUZZ_REPORT_ONLY=1 logs a generated run's failure marker
 * without failing the run; pinned flows still fail.
 */

import { createHash } from "node:crypto";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { after, test } from "node:test";

import {
  failureMarker,
  logFailureMarker,
  reportScenarioFailure,
  shellQuote,
} from "../support/failure-fingerprints.ts";
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
const SAVE_SAMPLE = integer(process.env["FOLIO_SCENARIO_SAVE_SAMPLE"], 0);
const SAMPLE_DIR = process.env["FOLIO_SCENARIO_SAMPLE_DIR"];
const REPORT_ONLY = process.env["FOLIO_SCENARIO_FUZZ_REPORT_ONLY"] === "1";

const saveSample = async (
  bytes: Uint8Array,
  { seed, kind, fixture, mode }: { seed: number; kind: FlowKind; fixture: string; mode: string },
): Promise<void> => {
  if (SAMPLE_DIR === undefined) {
    throw new Error("FOLIO_SCENARIO_SAMPLE_DIR is required to save samples");
  }
  const fingerprint = createHash("sha256")
    .update(`${kind}\0${seed}\0${String(STEPS)}\0${fixture}\0${mode}`)
    .digest("hex")
    .slice(0, 16);
  await mkdir(SAMPLE_DIR, { recursive: true });
  await writeFile(path.join(SAMPLE_DIR, `${fingerprint}-${seed}.docx`), bytes);
};

const relationRepro = [
  "FOLIO_SCENARIO_RELATIONS",
  "FOLIO_SCENARIO_RELATIONS_DEPTH",
  "FOLIO_SCENARIO_SWARM",
  "FOLIO_SCENARIO_FEATURE_WEIGHTS",
]
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
      const sampleIndex = kind === "random" ? run : RUNS + run;
      const shouldSave = sampleIndex < SAVE_SAMPLE;
      let saved: Uint8Array | undefined;
      try {
        ({ saved } = await runFlow(seed, STEPS, kind, { captureSaved: shouldSave }));
      } catch (error) {
        const failure = { test: `consumer flow ${fixture} / ${mode}`, seed, repro, failure: error };
        if (REPORT_ONLY) {
          logFailureMarker(failureMarker(failure));
          console.warn(`report-only: ${label} run ${run} failed; replay: ${repro}`);
          return;
        }
        reportScenarioFailure(failure);
      }
      // Outside the try: a sample that cannot be written fails the run.
      if (shouldSave && saved !== undefined) {
        await saveSample(saved, { seed, kind, fixture, mode });
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
