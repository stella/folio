/** Full-suite timing provenance comes from the plan the shard actually executed. */
import { panic } from "better-result";

import type { TestTimings } from "./ci-test-shard-plan";

export type FullSuiteTimingArtifact = {
  version: 1;
  origin: { type: "github"; runId: string } | { type: "local" };
  shard: number;
  shards: number;
  timings: TestTimings;
};

const assertShard = (shard: number, shards: number): void => {
  if (
    !Number.isSafeInteger(shards) ||
    shards < 1 ||
    !Number.isSafeInteger(shard) ||
    shard < 1 ||
    shard > shards
  )
    panic(
      "Full-suite timings require a valid executed shard and planned shard count",
    );
};

const assertRunId = (runId: string): void => {
  if (!/^[1-9]\d*$/u.test(runId))
    panic("Full-suite timings require a GitHub run id");
};

type CreateFullSuiteTimingArtifactOptions = {
  runId: string | undefined;
  shard: number;
  shards: number;
  timings: TestTimings;
};

export const createFullSuiteTimingArtifact = ({
  runId,
  shard,
  shards,
  timings,
}: CreateFullSuiteTimingArtifactOptions): FullSuiteTimingArtifact => {
  assertShard(shard, shards);
  if (runId === undefined)
    return { version: 1, origin: { type: "local" }, shard, shards, timings };
  assertRunId(runId);
  return {
    version: 1,
    origin: { type: "github", runId },
    shard,
    shards,
    timings,
  };
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Both timing producers report positive millisecond measurements keyed by file. */
export const readFileTimings = (
  source: unknown,
  runId: string,
): TestTimings => {
  if (!isRecord(source))
    return panic("Timing artifact is missing file measurements", { runId });
  const timings: TestTimings = {};
  for (const [file, duration] of Object.entries(source)) {
    if (
      typeof duration !== "number" ||
      !Number.isFinite(duration) ||
      duration < 1
    )
      return panic("File timings must be positive finite milliseconds", {
        runId,
        file,
      });
    timings[file] = duration;
  }
  return timings;
};

type ParsedArtifact =
  | { type: "legacy"; runId: string }
  | {
      type: "proven";
      runId: string;
      shard: number;
      shards: number;
      timings: TestTimings;
    };

const parseArtifact = (value: unknown): ParsedArtifact => {
  if (!isRecord(value))
    return panic("Full-suite timing artifact must be an object");
  if (
    value["version"] === undefined &&
    typeof value["runId"] === "string" &&
    value["shard"] === undefined &&
    value["shards"] === undefined &&
    value["origin"] === undefined
  ) {
    readFileTimings(value["timings"], value["runId"]);
    return { type: "legacy", runId: value["runId"] };
  }
  if (value["version"] !== 1)
    return panic("Unsupported full-suite timing artifact version");
  const origin = value["origin"];
  if (!isRecord(origin))
    return panic("Full-suite timings require an explicit origin");
  if (origin["type"] === "local")
    return panic("Local full-suite timings cannot refresh CI measurements");
  if (origin["type"] !== "github" || typeof origin["runId"] !== "string")
    return panic("Full-suite timings require a GitHub origin");
  const runId = origin["runId"];
  assertRunId(runId);
  const shard = value["shard"];
  const shards = value["shards"];
  if (typeof shard !== "number" || typeof shards !== "number")
    return panic("Full-suite timings are missing their executed shard facts");
  assertShard(shard, shards);
  const timings = readFileTimings(value["timings"], runId);
  return { type: "proven", runId, shard, shards, timings };
};

type FullSuiteTimingInput = {
  /** The run owning this artifact in GitHub's download manifest. */
  runId: string;
  artifact: unknown;
};

/**
 * Validate every selected run before returning any measurements. The pre-envelope
 * producer's artifacts expire after 14 days: remove legacy skipping once none of
 * those eligible runs remain within artifact retention. Their count is unknown,
 * so they are reported as unavailable, never guessed to be complete.
 */
export const fullSuiteTimingSets = (
  values: readonly FullSuiteTimingInput[],
) => {
  const runs = new Map<string, Extract<ParsedArtifact, { type: "proven" }>[]>();
  const legacyRuns = new Set<string>();
  for (const { runId, artifact } of values) {
    assertRunId(runId);
    const parsed = parseArtifact(artifact);
    if (parsed.runId !== runId)
      panic(
        "A timing artifact names a different run than its download manifest",
        { expected: runId, actual: parsed.runId },
      );
    switch (parsed.type) {
      case "legacy":
        legacyRuns.add(parsed.runId);
        break;
      case "proven": {
        const run = runs.get(parsed.runId) ?? [];
        run.push(parsed);
        runs.set(parsed.runId, run);
        break;
      }
      default: {
        const unreachable: never = parsed;
        return unreachable;
      }
    }
  }
  const artifacts: { runId: string; timings: TestTimings }[] = [];
  for (const [runId, run] of runs) {
    if (legacyRuns.has(runId))
      panic("One timing run mixes legacy and proven shard metadata", { runId });
    const shards =
      run.at(0)?.shards ?? panic("A timing run lost its artifacts");
    const seen = new Set<number>();
    const files = new Set<string>();
    for (const artifact of run) {
      if (artifact.shards !== shards)
        panic("One timing run declares inconsistent planned shard counts", {
          runId,
        });
      if (seen.has(artifact.shard))
        panic("A timing run repeats an executed shard", {
          runId,
          shard: artifact.shard,
        });
      seen.add(artifact.shard);
      for (const file of Object.keys(artifact.timings)) {
        if (files.has(file))
          panic("A timing run assigns one file to multiple shards", {
            runId,
            file,
          });
        files.add(file);
      }
    }
    if (seen.size !== shards)
      panic("A timing run is missing planned shard measurements", {
        runId,
        planned: shards,
        present: [...seen].toSorted((a, b) => a - b),
      });
    artifacts.push(...run.map(({ timings }) => ({ runId, timings })));
  }
  return { artifacts, legacyRuns: [...legacyRuns].toSorted() };
};
