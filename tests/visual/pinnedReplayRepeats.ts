import { createHash } from "node:crypto";
import path from "node:path";

/** Where a pinned replay writes what it observed, when CI repeats the replays. */
export const PINNED_REPLAY_OUTCOMES_ENV = "FOLIO_PINNED_REPLAY_OUTCOMES";

/**
 * The font delivery delay of each repeat: the first gets the fonts as the
 * server sends them, the later ones after the editor has painted, so a target
 * read from paint a late face moves is read differently in some repeat.
 */
export const PINNED_REPLAY_FONT_DELAYS_MS: readonly number[] = [0, 150, 600];

/** How many times CI replays each pinned trace: once per font timing. */
export const PINNED_REPLAY_REPEATS: number =
  PINNED_REPLAY_FONT_DELAYS_MS.length;

type OutcomeFileOptions = {
  directory: string;
  title: string;
  repeat: number;
};

export const pinnedReplayOutcomeFile = ({
  directory,
  title,
  repeat,
}: OutcomeFileOptions): string =>
  path.join(
    directory,
    `${createHash("sha256").update(title).digest("hex").slice(0, 16)}-${String(repeat)}.json`,
  );

export type PinnedReplayOutcome = {
  title: string;
  repeat: number;
  outcome: "passed" | "failed";
  observations: unknown;
};

/**
 * Every way the repeats of the pinned replays disagree: a replay with another
 * number of repeats than `repeats`, or a repeat whose outcome or observations
 * differ from its first. Empty when every replay observed the same each time.
 */
export const pinnedReplayDifferences = (
  outcomes: readonly PinnedReplayOutcome[],
  repeats: number,
): string[] => {
  const byTitle = new Map<string, PinnedReplayOutcome[]>();
  for (const outcome of outcomes) {
    const runs = byTitle.get(outcome.title) ?? [];
    runs.push(outcome);
    byTitle.set(outcome.title, runs);
  }
  const differences: string[] = [];
  for (const [title, runs] of [...byTitle].toSorted(([left], [right]) =>
    left.localeCompare(right),
  )) {
    const ordered = runs.toSorted((left, right) => left.repeat - right.repeat);
    const indices = ordered.map(({ repeat }) => repeat);
    if (
      indices.length !== repeats ||
      indices.some((repeat, index) => repeat !== index)
    ) {
      differences.push(
        `${title}: repeats ${JSON.stringify(indices)}, expected 0..${String(repeats - 1)}`,
      );
      continue;
    }
    const [first, ...rest] = ordered;
    if (!first) continue;
    const expected = JSON.stringify({
      outcome: first.outcome,
      observations: first.observations,
    });
    for (const run of rest) {
      const actual = JSON.stringify({
        outcome: run.outcome,
        observations: run.observations,
      });
      if (actual !== expected)
        differences.push(
          `${title}: repeat ${String(run.repeat)} observed ${actual}, repeat 0 observed ${expected}`,
        );
    }
  }
  return differences;
};
