/**
 * Scenarios that reproduce an open issue run as expected failures: each must
 * still fail, with an error matching the issue's symptom. When the fix lands
 * the scenario passes, the expected failure fails ("no longer reproduces"),
 * and the marker is removed so the scenario guards the fix from then on.
 */

import assert from "node:assert/strict";
import { test } from "node:test";

export const OPEN_ISSUES = {
  1091: "list commands reference a numbering instance the package does not define",
  1092: "list commands always join the first list of their kind",
  1094: "a numbered heading reads as listItem; read_document rows carry no label",
  1103: "operations accept a numbering.numId the package does not define",
} as const;

/**
 * Found by these scenarios and not yet filed or fixed; each has a minimal
 * repro in the scenario that pins it.
 */
export const FINDINGS = {
  STALE_LIST_LABELS:
    "after an operation adds, removes or renumbers list items, the reviewer's getContent() / snapshot labels keep the numbers read at open until a save and reopen",
  COMMENT_ANCHOR_DRIFT:
    "a comment whose range an edit splits or replaces reads a different anchoredText before the save (the marked text) than after it (everything between the range markers)",
  NOTE_REFERENCE_TEXT:
    'getContent(), the snapshot and read_document render a footnote/endnote reference as its w:id ("7"), while the page shows its number ("1", or "i" for an endnote)',
  COMPARE_INSERTED_LIST_ITEMS:
    "generateRedlineDocx inserts a paragraph the revised version has as a list item as a plain paragraph (the insertion carries only its style), so accepting the redline loses the bullet or number",
  REJECT_SPLIT_AROUND_INSERTED_TABLE:
    "rejecting every change leaves a tracked split in place when a tracked table was inserted after the split's first half (the join is attempted while the table still stands between the halves)",
  BATCH_SPLIT_THEN_DELETE:
    "a tracked batch that splits a block and deletes the same block deletes the wrong span: part of the text survives, and accepting leaves an empty paragraph",
} as const;

export type OpenIssue = keyof typeof OPEN_ISSUES;
export type Finding = keyof typeof FINDINGS;

const describeKnown = (known: OpenIssue | Finding): string =>
  typeof known === "number" ? `#${known} (${OPEN_ISSUES[known]})` : `${known} (${FINDINGS[known]})`;

/** Whether readers should tolerate an issue's known disagreement. */
export const TOLERATED_DISAGREEMENTS: ReadonlySet<OpenIssue> = new Set<OpenIssue>([1094]);

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
