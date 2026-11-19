/**
 * Every checked-in flow replayed with every reviewer on the canonical
 * session, through the packed public entry points: each step, save, reopen
 * and resolution is held to the same invariants the default replay holds,
 * and its steps must act as the default replay's do. The scenario files
 * `canonical-flows-*.test.ts` each take one shard of the flows, so the
 * replays run side by side.
 *
 * A flow the canonical session does not pass yet is listed in
 * {@link CANONICAL_REPLAY_GAPS} with why and how it fails; it must still fail
 * that way, so a fix that makes it pass removes it from the list and the
 * flow guards the fix.
 */

import assert from "node:assert/strict";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

import { readFlowDir } from "./corpus.ts";
import { scenarioEngine } from "./documents.ts";
import { runFlowFile, vacuousSteps } from "./fuzz.ts";
import { VACUOUS_CHECKED_IN_FLOWS } from "./known-issues.ts";

const FLOWS_DIR = fileURLToPath(new URL("../flows/", import.meta.url));

type ReplayGap = {
  reason: string;
  /** What the replay's failure says; any other failure is not this gap. */
  symptom: RegExp;
};

const MARKDOWN_ACCEPTANCE: ReplayGap = {
  reason: "clean Markdown still resolves with the default engine's acceptance",
  symptom: /Markdown contains more text blocks than the source reader/u,
};

const SEQUENTIAL_DELETION: ReplayGap = {
  reason:
    "the batch's block deletion beside the table it inserts is refused one operation at a time",
  symptom:
    /\[batchSequential\][^\n]*\n\s+op-2 refused one at a time: unsupportedBlock/u,
};

/** Checked-in flows the canonical session does not pass yet. */
export const CANONICAL_REPLAY_GAPS: Readonly<Record<string, ReplayGap>> = {
  "comment-replacement-point.json": {
    reason:
      "a commented block's deletion before the table the batch inserts after it is refused",
    symptom: /step 4: expected 2 applied operations, got 1/u,
  },
  "issue-1350-7fe54148e47f13fe.json": MARKDOWN_ACCEPTANCE,
  "issue-1350-table-and-paragraph-before-deleted-final.json":
    MARKDOWN_ACCEPTANCE,
  "issue-1356-3058fb47c2ef6a7c.json": SEQUENTIAL_DELETION,
  "issue-1356-table-after-a-deleted-paragraph-joining-the-next.json": {
    reason:
      "a table put after a wordless paragraph joining the next goes before a paragraph the batch put before that next one earlier",
    symptom: /the result is not what was asked[^\n]*\n\s+block texts differ/u,
  },
  "issue-1356-table-after-a-deleted-final-paragraph.json": SEQUENTIAL_DELETION,
  "issue-1757-7483271c36812e15.json": MARKDOWN_ACCEPTANCE,
  "issue-1757-paragraph-before-table-after-pending-join.json":
    MARKDOWN_ACCEPTANCE,
};

/**
 * Checked-in flows whose steps act on nothing on the canonical session where
 * the default replay's do not, and those steps; the rest act as
 * VACUOUS_CHECKED_IN_FLOWS states for the default replay.
 */
export const CANONICAL_VACUOUS_STEPS: Readonly<
  Record<string, readonly number[]>
> = {
  "accepted-tracked-replacement.json": [0],
  "issue-1350-e377f5b27fbc60fc.json": [1],
  "issue-1356-42bb7d148b3e1e55.json": [1, 2],
  "issue-1356-split-after-insertion-beside-pending-join.json": [4],
  "issue-1754-cdc101c598f26387.json": [0],
  "issue-1757-merge-across-deleted-note-paragraph-before-inserted-table.json": [
    0,
  ],
  "issue-1757-merge-across-deleted-paragraph-before-inserted-table.json": [0],
  "issue-1762-reject-deleted-first-row-after-a-split-half.json": [4],
  "issue-1811-comment-start-in-a-deleted-cell.json": [2],
  "issue-1842-comment-on-split-note-tail.json": [2],
  "issue-1842-comment-quotes-across-line-breaks.json": [0, 1],
  "issue-1842-final-deletion-retracts-split-break-after-reopen.json": [4],
  "issue-2044-heading-style-in-a-table-cell.json": [12],
  "issue-2222-styled-insert-after-undefined-numbering-level.json": [2],
  "issue-2236-new-list-in-header-after-other-stories-direct.json": [0],
  "issue-2316-c14cd1e28851babe.json": [0, 1],
  "issue-2904-merged-numbered-heading-emptied.json": [1],
  "markdown-note-reference-emphasis.json": [0, 1],
  "revision-boundary-1c6f520a41a4dcf4.json": [4],
  "revision-round-trip-1.json": [4],
  "suggested-table-row-beside-replaced-paragraph.json": [6, 7, 8, 12, 13],
  "undefined-numbering-level-projection.json": [1],
};

/** How many scenario files share the canonical replays. */
export const CANONICAL_REPLAY_SHARDS = 4;

/**
 * Register the replays of shard `shard` (0-based) of the checked-in flows,
 * in file order. The scenario file sets `FOLIO_SCENARIO_ENGINE` before it
 * opens anything, since each scenario file runs in its own process.
 */
export const registerCanonicalReplays = (shard: number): void => {
  test("the replays run on the canonical session", () => {
    assert.equal(scenarioEngine(), "canonical");
  });
  const corpus = readFlowDir(FLOWS_DIR).toSorted((left, right) =>
    left.file.localeCompare(right.file),
  );
  for (const [index, { file, flow }] of corpus.entries()) {
    if (index % CANONICAL_REPLAY_SHARDS !== shard) continue;
    const gap = CANONICAL_REPLAY_GAPS[file];
    if (gap === undefined) {
      test(`${file} replays on the canonical session`, async () => {
        const run = await runFlowFile(flow);
        assert.deepEqual(
          vacuousSteps(run),
          CANONICAL_VACUOUS_STEPS[file] ?? VACUOUS_CHECKED_IN_FLOWS[file] ?? [],
          `${file}: the steps that act on nothing on the canonical session are not the ones CANONICAL_VACUOUS_STEPS or VACUOUS_CHECKED_IN_FLOWS lists`,
        );
      });
      continue;
    }
    test(`${file} still fails on the canonical session (${gap.reason})`, async () => {
      await assert.rejects(
        runFlowFile(flow),
        (error) => error instanceof Error && gap.symptom.test(error.message),
        `${file} no longer fails as CANONICAL_REPLAY_GAPS records: remove its entry once it passes`,
      );
    });
  }
  if (shard !== 0) return;
  test("every canonical replay exception names a checked-in flow", () => {
    const files = new Set(readFlowDir(FLOWS_DIR).map(({ file }) => file));
    assert.deepEqual(
      [
        ...Object.keys(CANONICAL_REPLAY_GAPS),
        ...Object.keys(CANONICAL_VACUOUS_STEPS),
      ].filter((file) => !files.has(file)),
      [],
    );
  });
};
