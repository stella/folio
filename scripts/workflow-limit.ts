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

/**
 * The fallback until a step has a measurement, and a floor afterwards. A
 * measurement needing more than the reviewed maximum throws, so a slowdown
 * is raised by a change someone reads rather than by the refresh.
 */
export const derivedMinutes = (
  durations: WorkflowDurations,
  { workflow, job, step }: Target,
  fallback: number,
  max: number,
): number => {
  const measurement = durations[workflow]?.[job]?.steps?.[step];
  if (measurement === undefined) return fallback;
  const needed = minimumMinutes(measurement.p95Seconds);
  if (needed > max)
    throw new RangeError(
      `${step}: measured p95 ${(measurement.p95Seconds / 60).toFixed(1)} min needs a ${needed} min limit, which exceeds the reviewed max ${max}; raise it in ${workflow} if intended`,
    );
  return Math.max(fallback, needed);
};

if (import.meta.main) {
  const { values } = parseArgs({
    options: {
      workflow: { type: "string" },
      job: { type: "string" },
      step: { type: "string" },
      fallback: { type: "string" },
      max: { type: "string" },
    },
  });
  const fallback = Number(values.fallback);
  const max = Number(values.max);
  if (
    values.workflow === undefined ||
    values.job === undefined ||
    values.step === undefined ||
    !Number.isInteger(fallback) ||
    fallback <= 0 ||
    !Number.isInteger(max) ||
    max < fallback
  ) {
    console.error(
      "Usage: workflow-limit.ts --workflow <file> --job <id> --step <name> --fallback <minutes> --max <minutes>",
    );
    process.exit(2);
  }
  try {
    const minutes = derivedMinutes(
      readWorkflowDurations(),
      { workflow: values.workflow, job: values.job, step: values.step },
      fallback,
      max,
    );
    console.log(`minutes=${minutes}`);
  } catch (error) {
    console.error(`::error::${error instanceof Error ? error.message : error}`);
    process.exit(1);
  }
}
