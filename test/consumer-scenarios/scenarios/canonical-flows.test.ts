/**
 * Checked-in flows replayed with every reviewer on the canonical session,
 * through the packed public entry points: each step, save, reopen and
 * resolution is held to the same invariants the legacy replay holds.
 *
 * The flows here hold edits the two saves write apart, which each must pass
 * on both engines. Each scenario file runs in its own process, so the engine
 * set here reaches only this file's reviewers.
 */

import { test } from "node:test";
import { fileURLToPath } from "node:url";

import { readFlowDir } from "../support/corpus.ts";
import { scenarioEngine } from "../support/documents.ts";
import { runFlowFile } from "../support/fuzz.ts";

process.env["FOLIO_SCENARIO_ENGINE"] = "canonical";

const FLOWS_DIR = fileURLToPath(new URL("../flows/", import.meta.url));

/** The flows replayed on the canonical session as well. */
const CANONICAL_REPLAYS = [
  "issue-1757-merge-into-closing-paragraph-beside-a-pending-join.json",
  "issue-1757-merge-into-closing-paragraph-of-deleted-words.json",
  "issue-1350-48a4c26dd3e1e6fd.json",
  "issue-1811-comment-start-in-a-deleted-cell.json",
  "issue-1762-tracked-formatting-between-line-breaks.json",
  "issue-1356-table-after-a-deleted-final-paragraph-direct.json",
  "tracked-comment-point.json",
  "unicode-comment-point.json",
  "issue-1754-clear-numbered-heading-style.json",
  "issue-1350-new-list-after-deleted-list-paragraph.json",
  "issue-2265-final-mark-rotation-over-a-merge-record.json",
  "issue-2973-reject-joins-into-terminal-table-carrier.json",
  "markdown-orphan-roman-level.json",
  "issue-1757-paragraph-after-emptied-paragraph-pending-break.json",
  "issue-1356-58c8b2f17ad44ed2.json",
  "issue-1356-865c9c708580cbcc.json",
  "issue-1356-c4f902fd90676d63.json",
  "issue-1817-insert-before-a-pending-merge.json",
] as const;

test("the replays below run on the canonical session", () => {
  if (scenarioEngine() !== "canonical")
    throw new Error("canonical flow replays opened a legacy reviewer");
});

for (const { file, flow } of readFlowDir(FLOWS_DIR)) {
  if (!CANONICAL_REPLAYS.some((replayed) => replayed === file)) continue;
  test(`${file} replays on the canonical session`, async () => {
    await runFlowFile(flow);
  });
}

test("every flow replayed on the canonical session is checked in", () => {
  const files = new Set(readFlowDir(FLOWS_DIR).map(({ file }) => file));
  const missing = CANONICAL_REPLAYS.filter((file) => !files.has(file));
  if (missing.length > 0)
    throw new Error(`missing flow files: ${missing.join(", ")}`);
});
