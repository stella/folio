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
  TERMINAL_DELETE_BATCH_FORMATTING:
    "accepting a tracked batch with preceding paragraph formatting and terminal deletion loses the requested formatting",
  INSERT_AFTER_PENDING_MERGE:
    "insertAfterBlock on a block whose tracked merge with the next is pending lists the new paragraph between them, but accepting joins the new paragraph onto the merged block and leaves the block the merge named apart",
  MARKDOWN_DROPS_TEXT_BOX:
    "docxToMarkdown writes nothing of a text box's paragraphs, which getContent() and read_document list as blocks (support/readers.ts leaves them out of the Markdown comparison until fixed)",
  REJECT_KEEPS_PARAGRAPH_INSERTED_IN_DELETED_NOTE:
    "a tracked deletion of a footnote's reference paragraph, then a tracked paragraph inserted into that footnote: rejecting every change leaves the inserted paragraph behind, empty",
  MERGE_INTO_BLOCK_THE_BATCH_DELETES:
    "#1415: a batch merges a block with the next and deletes that next block; the legacy path applies the batch in reverse, so the merge joins the block after the deleted one, where the oracle expects the separator with nothing joined",
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
}[] = [
  { seed: 18568319, steps: 16, finding: "TERMINAL_DELETE_BATCH_FORMATTING" },
  { seed: 18568230, steps: 16, finding: "MERGE_INTO_BLOCK_THE_BATCH_DELETES" },
];

/** How each finding fails a scenario, so an expected failure fails for that reason only. */
export const FINDING_SYMPTOMS: Record<Finding, RegExp> = {
  TERMINAL_DELETE_BATCH_FORMATTING: /directAlignment is undefined, expected "center"/u,
  MARKDOWN_DROPS_TEXT_BOX: /docxToMarkdown writes no text-box paragraph/u,
  INSERT_AFTER_PENDING_MERGE: /accepting glues the inserted paragraph onto the merged one/u,
  REJECT_KEEPS_PARAGRAPH_INSERTED_IN_DELETED_NOTE:
    /\[rejectAll\] the flow's batches replayed tracked and rejected do not give the fixture back[^\n]*\n\s*\{"type":"footnote","noteId":\d+\}: \.blocks\[\d+\]: undefined → \{"kind":"heading","text":""/u,
  // A batch with both operations, whose expected block ends in the merge's separator.
  MERGE_INTO_BLOCK_THE_BATCH_DELETES:
    /not what was asked \((?=[^)]*mergeBlockWithNext)(?=[^)]*deleteBlock)[^)]*\):\n\s*block texts differ:\n\s*expected \[[^\n]* ","/u,
};

/**
 * Checked-in flow files (scenarios/flow-corpus.test.ts) that reproduce a
 * finding; they run as expected failures there.
 */
export const KNOWN_FAILING_CHECKED_IN_FLOWS: Readonly<Record<string, Finding>> = {
  "merge-with-next-block-the-batch-deletes.json": "MERGE_INTO_BLOCK_THE_BATCH_DELETES",
  "suggested-note-delete-and-heading-insert.json":
    "REJECT_KEEPS_PARAGRAPH_INSERTED_IN_DELETED_NOTE",
};

/**
 * Checked-in flow files with steps that apply nothing (support/fuzz.ts
 * `vacuousSteps`), and those steps. Their pinned block ids are paragraph ids
 * minted from a hash of the fixture's document.xml as it serialized when
 * they were recorded; they name no block now, so these steps guard nothing.
 * Each replays exactly this vacuous until its ids are re-pinned from a
 * fresh shrink, when its entry goes.
 */
export const VACUOUS_CHECKED_IN_FLOWS: Readonly<Record<string, readonly number[]>> = {
  "batch-after-table-row-delete.json": [0, 1, 4, 5],
  "deleted-endnote-reopen.json": [0],
  "revision-wrappers-around-link.json": [0],
  "suggested-story-insert-and-delete.json": [0],
  "text-box-comment-reference-order.json": [2],
  "tracked-story-row-delete-and-column-insert.json": [0, 1],
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
