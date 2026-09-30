/**
 * Scenarios that reproduce an open issue run as expected failures: each must
 * still fail, with an error matching the issue's symptom. When the fix lands
 * the scenario passes, the expected failure fails ("no longer reproduces"),
 * and the marker is removed so the scenario guards the fix from then on.
 */

import assert from "node:assert/strict";
import { test } from "node:test";
import type { Relation } from "./metamorphic.ts";

export const OPEN_ISSUES = {} as const;

/**
 * Found by these scenarios and not yet filed or fixed; each has a minimal
 * repro in the scenario that pins it.
 */
export const FINDINGS = {
  SUGGESTED_NOTE_EDIT_SAVE_THROWS:
    "after a save that dropped a suggested format change in a footnote, a second suggested edit of that footnote plus any suggested body edit makes toBuffer() throw",
  INSERT_AFTER_PENDING_MERGE:
    "insertAfterBlock on a block whose tracked merge with the next is pending lists the new paragraph between them, but accepting joins the new paragraph onto the merged block and leaves the block the merge named apart",
  MARKDOWN_DROPS_TEXT_BOX:
    "docxToMarkdown writes nothing of a text box's paragraphs, which getContent() and read_document list as blocks (support/readers.ts leaves them out of the Markdown comparison until fixed)",
} as const;

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
export const KNOWN_FAILING_FLOWS: readonly {
  seed: number;
  steps: number;
  finding: Finding;
  /** The flow kind (support/fuzz.ts); `"random"` when absent. */
  kind?: "random" | "collisions";
  /** The generation (support/fuzz.ts); `"targeted"` when absent. */
  generation?: "targeted" | "legacy";
  /** Required relation for a finding; the scenario is omitted when disabled. */
  relation?: Relation;
}[] = [];

/** How each finding fails a scenario, so an expected failure fails for that reason only. */
export const FINDING_SYMPTOMS: Record<Finding, RegExp> = {
  SUGGESTED_NOTE_EDIT_SAVE_THROWS: /Cannot serialize changed footnote paragraphs/u,
  MARKDOWN_DROPS_TEXT_BOX: /docxToMarkdown writes no text-box paragraph/u,
  INSERT_AFTER_PENDING_MERGE: /accepting glues the inserted paragraph onto the merged one/u,
};

/**
 * requested-outcome.test.ts collisions (fixture / mode / collision) that
 * reach a finding; they run as expected failures there.
 */
export const KNOWN_FAILING_COLLISIONS: readonly {
  fixture: string;
  mode: string;
  collision: string;
  finding: Finding;
  symptom: RegExp;
}[] = [];

/** requested-outcome.test.ts follow-up edits of pending changes that reach a finding. */
export const KNOWN_FAILING_FOLLOW_UPS: readonly {
  followUp: string;
  mode: string;
  finding: Finding;
  symptom: RegExp;
}[] = [];

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
