/** Heavy flows are explicitly enabled on CI; the ordinary consumer suite skips them. */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { test } from "node:test";

import { writeFailureRecord } from "../support/failure-fingerprints.ts";
import { recordFailure, relationEnv } from "../support/fuzz-loop.ts";
import { FlowError, runFlow } from "../support/fuzz.ts";
import { LARGE_DOCUMENT_FIXTURE } from "../support/large-document.ts";

const ENABLED = process.env["CI"] === "true" && process.env["FOLIO_SCENARIO_LONG_FUZZ"] === "1";

if (!ENABLED) {
  test(
    "long flows on a 100-page document",
    { skip: "CI-only: FOLIO_SCENARIO_LONG_FUZZ=1" },
    () => {},
  );
} else {
  const seed = Number(process.env["FOLIO_SCENARIO_SEED"]);
  const steps = Number(process.env["FOLIO_SCENARIO_LONG_STEPS"]);
  const output = process.env["FOLIO_SCENARIO_LONG_OUTPUT"];
  if (
    !Number.isSafeInteger(seed) ||
    seed < 0 ||
    !Number.isInteger(steps) ||
    steps < 200 ||
    steps > 1_000 ||
    !output
  ) {
    throw new TypeError(
      "long fuzz requires a seed, 200–1000 steps, and FOLIO_SCENARIO_LONG_OUTPUT",
    );
  }
  test(`long flow: seed ${seed}, ${steps} steps, 100 pages`, async () => {
    await mkdir(output, { recursive: true });
    const repro = [
      `CI=true FOLIO_SCENARIO_LONG_FUZZ=1 FOLIO_SCENARIO_SEED=${seed} FOLIO_SCENARIO_LONG_STEPS=${steps} FOLIO_SCENARIO_LONG_OUTPUT=test-results/long-fuzz`,
      relationEnv(),
      "bun scripts/consumer-scenarios.ts -- long-fuzz.test.ts",
    ]
      .filter(Boolean)
      .join(" ");
    await writeFile(path.join(output, "replay.txt"), `${repro}\n`);
    try {
      const result = await runFlow(seed, steps, "collisions", {
        fixture: LARGE_DOCUMENT_FIXTURE,
      });
      await writeFile(path.join(output, "completed-flow.json"), `${JSON.stringify(result.flow)}\n`);
      console.log(`Completed ${result.flow.steps.length} steps; recorded completed-flow.json`);
    } catch (error) {
      if (!(error instanceof FlowError)) throw error;
      await writeFile(path.join(output, "original-flow.json"), `${JSON.stringify(error.flow)}\n`);
      const record = await recordFailure(error, { seed, repro }, { maxAttempts: 40, seconds: 240 });
      writeFailureRecord(output, record);
      throw error;
    }
  });
}
