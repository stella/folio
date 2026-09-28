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
  LIVE_COMMENT_ANCHOR_STALE:
    "a paragraph inserted inside a comment range that spans several blocks is part of the anchored text after a save and reopen, but not in the live reviewer's getComments()",
  TEXT_BOX_RESOLVE_MALFORMED_XML:
    'with tracked changes pending both in a paragraph that draws a text box and in the text box, acceptAll() or rejectAll() after a reopen saves a document.xml that is not well-formed ("</w:p>></w:r>") and does not reopen',
  TRACKED_MERGE_INTO_DELETED_BLOCK:
    "a tracked-changes batch that deletes a block and merges the block before it into it applies both, where direct mode refuses the merge as overlappingOperation; accepting then joins the merged block to the block after the deleted one (or leaves the separator dangling at the end)",
  INSERT_AFTER_PENDING_MERGE:
    "insertAfterBlock on a block whose tracked merge with the next is pending lists the new paragraph between them, but accepting joins the new paragraph onto the merged block and leaves the block the merge named apart",
  MARKDOWN_DROPS_TEXT_BOX:
    "docxToMarkdown writes nothing of a text box's paragraphs, which getContent() and read_document list as blocks (support/readers.ts leaves them out of the Markdown comparison until fixed)",
  // Found by the metamorphic relations (support/metamorphic.ts) and the
  // flows they run in. A relation tolerates the frequent ones through
  // `tolerate(<entry>, …)`, so deleting a fixed entry here makes its
  // tolerance fail to compile until it is deleted too.
  LIVE_STALE_BLOCK_FIELDS:
    "the live reviewer does not re-resolve what a paragraph's style gives it after an edit: a restyled paragraph keeps its old style's previewRuns, a paragraph inserted after a bold heading previews bold despite its direct bold off, and a paragraph inserted with a numbered heading style has no directIndentation; the saved package reopens with other values",
  LIVE_REPLY_RANGES:
    "a reply added with replyTo has no comment range in the live document, while the saved package anchors it on its parent's range, so toMarkdown(toDocument()) reads otherwise across a save",
  TRACKED_DELETE_LAST_PARAGRAPH:
    "deleteBlock of the story's last paragraph: applied directly it removes the paragraph, tracked and accepted it leaves an empty paragraph",
  TRACKED_LAST_PARAGRAPH_REJECT:
    "one tracked batch that deletes the story's last paragraph and inserts a paragraph after it: rejecting every change leaves an extra empty paragraph (a list item when the insertion started a list)",
  COMMENT_ANCHOR_REPLACED_BLOCK:
    "two comments on one paragraph, then replaceBlock: applied directly both comments anchor on the new text, tracked and accepted the first one's anchor is left empty",
  BATCH_ROW_INSERT_DELETE:
    "one direct batch that inserts a row after a table row and deletes the row below it: both report applied, but the row to delete stays",
  SAVE_REORDERS_COMMENT_RANGES:
    "after replies across several steps, saving the reopened package writes co-located commentRangeStart elements in another order than the save it was opened from",
  LIVE_GET_CHANGES_STALE:
    "after a legacy collision flow, getChanges() reads different change text, locations, or kinds before and after saving and reopening",
  TRACKED_TABLE_AFTER_SPLIT_DELETE:
    "a tracked split followed by deleting the new block and inserting a table differs from the equivalent direct edits after accepting changes",
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
  // From a 300-flow sweep (FOLIO_SCENARIO_SEED=7310000, 100 collision runs).
  {
    seed: 7_310_028,
    steps: 10,
    kind: "collisions",
    generation: "legacy",
    finding: "SAVE_REORDERS_COMMENT_RANGES",
    relation: "saveIdempotent",
  },
  {
    seed: 7_310_042,
    steps: 10,
    kind: "collisions",
    generation: "legacy",
    finding: "BATCH_ROW_INSERT_DELETE",
  },
  {
    seed: 1088,
    steps: 10,
    kind: "collisions",
    generation: "legacy",
    finding: "LIVE_GET_CHANGES_STALE",
    relation: "readerStability",
  },
  {
    seed: 1185,
    steps: 10,
    kind: "collisions",
    generation: "legacy",
    finding: "LIVE_GET_CHANGES_STALE",
    relation: "readerStability",
  },
  {
    seed: 20_260_937,
    steps: 10,
    finding: "TRACKED_TABLE_AFTER_SPLIT_DELETE",
    relation: "directTracked",
  },
];

/** How each finding fails a scenario, so an expected failure fails for that reason only. */
export const FINDING_SYMPTOMS: Record<Finding, RegExp> = {
  SUGGESTED_NOTE_EDIT_SAVE_THROWS: /Cannot serialize changed footnote paragraphs/u,
  LIVE_COMMENT_ANCHOR_STALE: /shows something else than the reviewer[\s\S]*\+\s+anchor: /u,
  MARKDOWN_DROPS_TEXT_BOX: /docxToMarkdown writes no text-box paragraph/u,
  TRACKED_MERGE_INTO_DELETED_BLOCK: /applied a merge into a block the batch deletes/u,
  INSERT_AFTER_PENDING_MERGE: /accepting glues the inserted paragraph onto the merged one/u,
  TEXT_BOX_RESOLVE_MALFORMED_XML: /malformed markup/u,
  LIVE_STALE_BLOCK_FIELDS: /previewRuns|directIndentation/u,
  LIVE_REPLY_RANGES: /\[readerStability\] toMarkdown/u,
  TRACKED_DELETE_LAST_PARAGRAPH: /\[directTracked\]/u,
  TRACKED_LAST_PARAGRAPH_REJECT: /\[rejectAll\]/u,
  COMMENT_ANCHOR_REPLACED_BLOCK: /\[directTracked\]/u,
  BATCH_ROW_INSERT_DELETE: /not what was asked \(deleteTableRow, insertTableRow\)/u,
  SAVE_REORDERS_COMMENT_RANGES: /\[saveIdempotent\][^\n]*\n?[^\n]*commentRange/u,
  LIVE_GET_CHANGES_STALE: /\[readerStability\] getChanges/u,
  TRACKED_TABLE_AFTER_SPLIT_DELETE: /\[directTracked\]/u,
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
