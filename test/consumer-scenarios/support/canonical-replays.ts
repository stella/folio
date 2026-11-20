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

/** Checked-in flows the canonical session does not pass yet. */
export const CANONICAL_REPLAY_GAPS: Readonly<Record<string, ReplayGap>> = {
  "issue-1350-7fe54148e47f13fe.json": MARKDOWN_ACCEPTANCE,
  "issue-1350-table-and-paragraph-before-deleted-final.json":
    MARKDOWN_ACCEPTANCE,
  "issue-1356-48eb3fcfbb6a954d.json": {
    reason:
      "a tracked insertion after a paragraph that would take over the last paragraph's pending property change is refused, which the direct replay applies",
    symptom:
      /\[directTracked\] the flow's batches replayed tracked and accepted read otherwise than replayed direct/u,
  },
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
  "issue-1356-42bb7d148b3e1e55.json": [1, 2],
  "issue-1356-split-after-insertion-beside-pending-join.json": [4],
  "issue-1757-merge-across-deleted-note-paragraph-before-inserted-table.json": [
    0,
  ],
  "issue-1757-merge-across-deleted-paragraph-before-inserted-table.json": [0],
  "issue-2044-heading-style-in-a-table-cell.json": [12],
  "issue-2222-styled-insert-after-undefined-numbering-level.json": [2],
  "issue-2236-new-list-in-header-after-other-stories-direct.json": [0],
  "issue-2316-c14cd1e28851babe.json": [0, 1],
  "issue-2904-merged-numbered-heading-emptied.json": [1],
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
