#!/usr/bin/env bun
// Prints `minutes=<n>` for a step's `timeout-minutes`: the smallest whole
// number of minutes that keeps the share of headroom over the step's measured
// p95 (scripts/workflow-durations.json), never under the fallback. A step
// reads it as `timeout-minutes: ${{ fromJSON(steps.<id>.outputs.minutes) }}`,
// so a timing refresh changes only the measurements.

import { parseArgs } from "node:util";

import {
  minimumMinutes,
  readWorkflowDurations,
  type WorkflowDurations,
} from "./workflow-durations";

type Target = { workflow: string; job: string; step: string };

/** The fallback until a step has a measurement, and a floor afterwards. */
export const derivedMinutes = (
  durations: WorkflowDurations,
  { workflow, job, step }: Target,
  fallback: number,
): number => {
  const measurement = durations[workflow]?.[job]?.steps?.[step];
  return measurement === undefined
    ? fallback
    : Math.max(fallback, minimumMinutes(measurement.p95Seconds));
};

if (import.meta.main) {
  const { values } = parseArgs({
    options: {
      workflow: { type: "string" },
      job: { type: "string" },
      step: { type: "string" },
      fallback: { type: "string" },
    },
  });
  const fallback = Number(values.fallback);
  if (
    values.workflow === undefined ||
    values.job === undefined ||
    values.step === undefined ||
    !Number.isInteger(fallback) ||
    fallback <= 0
  ) {
    console.error(
      "Usage: workflow-limit.ts --workflow <file> --job <id> --step <name> --fallback <minutes>",
    );
    process.exit(2);
  }
  const minutes = derivedMinutes(
    readWorkflowDurations(),
    { workflow: values.workflow, job: values.job, step: values.step },
    fallback,
  );
  console.log(`minutes=${minutes}`);
}
