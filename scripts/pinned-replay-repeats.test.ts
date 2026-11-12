import { expect, test } from "bun:test";

import {
  PINNED_REPLAY_FONT_DELAYS_MS,
  PINNED_REPLAY_REPEATS,
  pinnedReplayDifferences,
  type PinnedReplayOutcome,
  pinnedReplayOutcomeFile,
} from "../tests/visual/pinnedReplayRepeats";

const run = (
  title: string,
  repeat: number,
  observations: unknown,
  outcome: PinnedReplayOutcome["outcome"] = "passed",
): PinnedReplayOutcome => ({ title, repeat, outcome, observations });

const drag = (from: number) => [{ drag: "note", selected: { from, to: 63 } }];

test("every repeat runs under its own font timing, the first undelayed", () => {
  expect(PINNED_REPLAY_REPEATS).toBe(PINNED_REPLAY_FONT_DELAYS_MS.length);
  expect(PINNED_REPLAY_REPEATS).toBeGreaterThan(1);
  expect(PINNED_REPLAY_FONT_DELAYS_MS.at(0)).toBe(0);
  expect(new Set(PINNED_REPLAY_FONT_DELAYS_MS).size).toBe(
    PINNED_REPLAY_REPEATS,
  );
});

test("replays that observe the same in every repeat agree", () => {
  const outcomes = [0, 1, 2].flatMap((repeat) => [
    run("a", repeat, drag(44)),
    run("b", repeat, [{ blocks: [] }]),
  ]);
  expect(pinnedReplayDifferences(outcomes, 3)).toEqual([]);
});

test("a repeat that observes another selection or outcome is a difference", () => {
  expect(
    pinnedReplayDifferences(
      [run("a", 0, drag(44)), run("a", 1, drag(45)), run("a", 2, drag(44))],
      3,
    ),
  ).toEqual([
    `a: repeat 1 observed ${JSON.stringify({ outcome: "passed", observations: drag(45) })}, repeat 0 observed ${JSON.stringify({ outcome: "passed", observations: drag(44) })}`,
  ]);
  expect(
    pinnedReplayDifferences(
      [
        run("a", 0, drag(44)),
        run("a", 1, drag(44), "failed"),
        run("a", 2, drag(44)),
      ],
      3,
    ),
  ).toHaveLength(1);
});

test("a replay missing a repeat, or repeated twice, is a difference", () => {
  expect(
    pinnedReplayDifferences([run("a", 0, []), run("a", 2, [])], 3),
  ).toEqual(["a: repeats [0,2], expected 0..2"]);
  expect(
    pinnedReplayDifferences(
      [run("a", 0, []), run("a", 1, []), run("a", 1, []), run("a", 2, [])],
      3,
    ),
  ).toEqual(["a: repeats [0,1,1,2], expected 0..2"]);
});

test("each replay's repeats write distinct files", () => {
  const files = new Set(
    ["regression seed 29 path 15:2 / 8d01cc728c151a9c", "other"].flatMap(
      (title) =>
        [0, 1, 2].map((repeat) =>
          pinnedReplayOutcomeFile({ directory: "out", title, repeat }),
        ),
    ),
  );
  expect(files.size).toBe(6);
});
