/**
 * Checked-in flows replayed with every reviewer on the canonical session,
 * through the packed public entry points: each step, save, reopen and
 * resolution is held to the same invariants the legacy replay holds.
 *
 * The flows here are the ones a legacy save is known to fail and the
 * canonical save is expected to pass. Each scenario file runs in its own
 * process, so the engine set here reaches only this file's reviewers.
 */

import { test } from "node:test";
import { fileURLToPath } from "node:url";

import { readFlowDir } from "../support/corpus.ts";
import { scenarioEngine } from "../support/documents.ts";
import { runFlowFile } from "../support/fuzz.ts";
import { KNOWN_FAILING_CHECKED_IN_FLOWS } from "../support/known-issues.ts";

process.env["FOLIO_SCENARIO_ENGINE"] = "canonical";

const FLOWS_DIR = fileURLToPath(new URL("../flows/", import.meta.url));

/** The legacy save's known failures the canonical save retires. */
const RETIRED_ON_CANONICAL = [
  "issue-1762-tracked-formatting-between-line-breaks.json",
] as const;

test("the replays below run on the canonical session", () => {
  if (scenarioEngine() !== "canonical")
    throw new Error("canonical flow replays opened a legacy reviewer");
});

for (const { file, flow } of readFlowDir(FLOWS_DIR)) {
  if (!RETIRED_ON_CANONICAL.some((retired) => retired === file)) continue;
  test(`${file} replays on the canonical session (a known ${KNOWN_FAILING_CHECKED_IN_FLOWS[file] ?? "legacy"} failure)`, async () => {
    await runFlowFile(flow);
  });
}

test("every flow retired on the canonical session is checked in", () => {
  const files = new Set(readFlowDir(FLOWS_DIR).map(({ file }) => file));
  const missing = RETIRED_ON_CANONICAL.filter((file) => !files.has(file));
  if (missing.length > 0)
    throw new Error(`missing flow files: ${missing.join(", ")}`);
});
