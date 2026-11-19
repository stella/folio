/**
 * Every checked-in flow replayed with every reviewer on the canonical
 * session, through the packed public entry points: each step, save, reopen
 * and resolution is held to the same invariants the default replay holds.
 * The scenario files `canonical-flows-*.test.ts` each take one shard of the
 * flows, so the replays run side by side.
 *
 * A flow the canonical session does not pass yet is listed in
 * {@link CANONICAL_REPLAY_GAPS} with why; it must still fail, so a fix that
 * makes it pass removes it from the list and the flow guards the fix.
 */

import assert from "node:assert/strict";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

import { readFlowDir } from "./corpus.ts";
import { scenarioEngine } from "./documents.ts";
import { runFlowFile } from "./fuzz.ts";

const FLOWS_DIR = fileURLToPath(new URL("../flows/", import.meta.url));

/** Checked-in flows the canonical session does not pass yet, and why. */
export const CANONICAL_REPLAY_GAPS: Readonly<Record<string, string>> = {
  "comment-replacement-point.json":
    "a block deletion beside a table the batch inserts is refused one block at a time",
  "issue-1350-7fe54148e47f13fe.json":
    "clean Markdown still resolves with the default engine's acceptance",
  "issue-1350-table-and-paragraph-before-deleted-final.json":
    "clean Markdown still resolves with the default engine's acceptance",
  "issue-1356-3058fb47c2ef6a7c.json":
    "the expectation keeps a deleted paragraph's words out of the following table's first cell, where accepting runs them",
  "issue-1356-table-after-a-deleted-final-paragraph.json":
    "the batch leaves otherwise than its operations one at a time",
  "issue-1757-7483271c36812e15.json":
    "clean Markdown still resolves with the default engine's acceptance",
  "issue-1757-paragraph-before-table-after-pending-join.json":
    "clean Markdown still resolves with the default engine's acceptance",
  "tracked-replacement-of-accepted-join-comments.json":
    "the replacement refuses for comment alignment, not for Reject All",
  "tracked-replacement-over-separate-comments.json":
    "the replacement refuses for comment alignment, not for Reject All",
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
        await runFlowFile(flow);
      });
      continue;
    }
    test(`${file} still fails on the canonical session (${gap})`, async () => {
      await assert.rejects(
        runFlowFile(flow),
        (error) => error instanceof Error,
        `${file} now passes on the canonical session: remove it from CANONICAL_REPLAY_GAPS`,
      );
    });
  }
  if (shard !== 0) return;
  test("every canonical replay gap names a checked-in flow", () => {
    const files = new Set(readFlowDir(FLOWS_DIR).map(({ file }) => file));
    assert.deepEqual(
      Object.keys(CANONICAL_REPLAY_GAPS).filter((file) => !files.has(file)),
      [],
    );
  });
};
