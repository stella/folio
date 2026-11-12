#!/usr/bin/env bun
// Compare what each pinned browser replay observed across its CI repeats:
// `bun scripts/pinned-replay-repeats.ts <outcome directory>`. Exits non-zero
// when a replay observed something else in any repeat, or ran another number
// of times than every font timing asks for.

import { readdir, readFile } from "node:fs/promises";
import path from "node:path";

import {
  PINNED_REPLAY_REPEATS,
  pinnedReplayDifferences,
  type PinnedReplayOutcome,
} from "../tests/visual/pinnedReplayRepeats";

const isOutcome = (value: unknown): value is PinnedReplayOutcome =>
  typeof value === "object" &&
  value !== null &&
  "title" in value &&
  typeof value.title === "string" &&
  "repeat" in value &&
  typeof value.repeat === "number" &&
  "outcome" in value &&
  (value.outcome === "passed" || value.outcome === "failed") &&
  "observations" in value;

const directory = process.argv[2];
if (directory === undefined) {
  console.error("Usage: bun scripts/pinned-replay-repeats.ts <directory>");
  process.exit(2);
}
const files = (await readdir(directory)).filter((file) =>
  file.endsWith(".json"),
);
if (files.length === 0) {
  console.error(`No pinned replay outcomes in ${directory}.`);
  process.exit(1);
}
const outcomes: PinnedReplayOutcome[] = [];
for (const file of files) {
  const value: unknown = JSON.parse(
    await readFile(path.join(directory, file), "utf8"),
  );
  if (!isOutcome(value)) {
    console.error(`${file} is not a pinned replay outcome.`);
    process.exit(1);
  }
  outcomes.push(value);
}
const differences = pinnedReplayDifferences(outcomes, PINNED_REPLAY_REPEATS);
if (differences.length > 0) {
  console.error(differences.join("\n"));
  process.exit(1);
}
console.log(
  `${String(outcomes.length / PINNED_REPLAY_REPEATS)} pinned replays observed the same in ${String(PINNED_REPLAY_REPEATS)} repeats.`,
);
