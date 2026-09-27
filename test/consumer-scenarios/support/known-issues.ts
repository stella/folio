/**
 * Scenarios that reproduce an open issue run as expected failures: each must
 * still fail, with an error matching the issue's symptom. When the fix lands
 * the scenario passes, the expected failure fails ("no longer reproduces"),
 * and the marker is removed so the scenario guards the fix from then on.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

export const OPEN_ISSUES = {
  1103: "operations accept a numbering.numId the package does not define",
} as const;

/**
 * Found by these scenarios and not yet filed or fixed; each has a minimal
 * repro in the scenario that pins it.
 */
export const FINDINGS = {} as const;

export type OpenIssue = keyof typeof OPEN_ISSUES;
export type Finding = keyof typeof FINDINGS;

const describeKnown = (known: OpenIssue | Finding): string =>
  typeof known === "number" ? `#${known} (${OPEN_ISSUES[known]})` : `${known} (${FINDINGS[known]})`;

/**
 * operations.test.ts runs (fixture / mode) whose operation sequence reaches a
 * finding; they run as expected failures there.
 */
export const KNOWN_FAILING_OPERATION_RUNS: readonly {
  fixture: string;
  mode: string;
  finding: Finding;
}[] = [];

/**
 * Seeded flows (support/fuzz.ts) that reproduce a finding. The default fuzz
 * run skips them and known-issues.test.ts runs them as expected failures.
 */
export const KNOWN_FAILING_FLOWS: readonly { seed: number; steps: number; finding: Finding }[] = [];

export const expectedFailure = (
  issue: OpenIssue | Finding,
  name: string,
  symptom: RegExp,
  body: () => Promise<void> | void,
): void => {
  const tag = typeof issue === "number" ? `#${issue}` : issue;
  test(`${name} [expected failure: ${tag}]`, async () => {
    let failure: unknown;
    try {
      await body();
    } catch (error) {
      failure = error;
    }
    if (failure === undefined) {
      assert.fail(
        `${describeKnown(issue)} no longer reproduces: "${name}" passes. ` +
          "Make it a plain scenario and drop the entry from support/known-issues.ts.",
      );
    }
    const message = failure instanceof Error ? failure.message : String(failure);
    if (!symptom.test(message)) {
      // It fails, but not the way the issue does: a real regression.
      throw failure;
    }
  });
};
