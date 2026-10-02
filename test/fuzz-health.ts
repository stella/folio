import type { RunDetails } from "fast-check";

import type { FuzzHealth } from "./consumer-scenarios/support/fuzz-health";

/** fast-check 4 exposes errorInstance; error is absent. */
export const fuzzErrorText = (error: unknown): string =>
  error instanceof Error ? (error.stack ?? error.message) : String(error);

export const classifyFuzzRun = <Ts>(details: RunDetails<Ts>): FuzzHealth => {
  if (details.failed) {
    const detail = fuzzErrorText(details.errorInstance);
    if (
      Array.isArray(details.counterexample) &&
      details.counterexample.length > 0 &&
      details.errorInstance !== undefined &&
      details.errorInstance !== null &&
      detail.trim() !== "" &&
      detail.trim() !== "undefined" &&
      detail.trim() !== "null"
    ) {
      return { status: "finding", completed: details.numRuns, detail };
    }
    return {
      status: "infrastructure",
      completed: details.numRuns,
      detail: "No valid counterexample or failure report",
    };
  }
  if (details.numRuns === 0 || details.interrupted) {
    return {
      status: "infrastructure",
      completed: details.numRuns,
      detail: "Fuzz run interrupted or executed zero cases",
    };
  }
  return { status: "passed", completed: details.numRuns };
};
