import type { RunDetails } from "fast-check";

import type { FuzzHealth } from "./consumer-scenarios/support/fuzz-health";

/** fast-check 4 exposes errorInstance; error is absent. */
export const fuzzErrorText = (error: unknown): string =>
  error instanceof Error ? (error.stack ?? error.message) : String(error);

export const classifyFuzzRun = <Ts>(details: RunDetails<Ts>): FuzzHealth => {
  if (
    details.numRuns === 0 ||
    details.interrupted ||
    (details.failed && details.counterexample === null)
  ) {
    return {
      status: "infrastructure",
      completed: details.numRuns,
      detail: "Fuzz run interrupted, exhausted skips, or executed zero cases",
    };
  }
  if (!details.failed) return { status: "passed", completed: details.numRuns };
  const detail = fuzzErrorText(details.errorInstance);
  if (
    !Array.isArray(details.counterexample) ||
    details.counterexample.length === 0 ||
    details.errorInstance === undefined ||
    detail.trim() === ""
  ) {
    return {
      status: "infrastructure",
      completed: details.numRuns,
      detail: "Empty counterexample or failure report",
    };
  }
  return { status: "finding", completed: details.numRuns, detail };
};
