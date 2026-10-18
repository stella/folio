const CONSUMER_SCENARIO_EXECUTABLE = "bun scripts/consumer-scenarios.ts";

type ReplayCommandOptions = { flags?: readonly string[]; files?: readonly string[] };

/** Build a replay with the executable first, followed by flags and file selectors. */
export const consumerScenarioReplay = ({
  flags = [],
  files = [],
}: ReplayCommandOptions = {}): string =>
  [
    CONSUMER_SCENARIO_EXECUTABLE,
    ...flags.filter((flag) => flag !== ""),
    ...(files.length === 0 ? [] : ["--", ...files]),
  ].join(" ");

export const CONSUMER_SCENARIO_REPLAY_INPUTS = [
  { flag: "--seed", env: "FOLIO_SCENARIO_SEED", kind: "value" },
  { flag: "--steps", env: "FOLIO_SCENARIO_FUZZ_STEPS", kind: "value" },
  { flag: "--runs", env: "FOLIO_SCENARIO_FUZZ_RUNS", kind: "value" },
  { flag: "--collision-runs", env: "FOLIO_SCENARIO_COLLISION_RUNS", kind: "value" },
  { flag: "--swarm", env: "FOLIO_SCENARIO_SWARM", kind: "value" },
  { flag: "--flow", env: "FOLIO_SCENARIO_FLOW", kind: "value" },
  { flag: "--public-corpus-dir", env: "FOLIO_SCENARIO_PUBLIC_CORPUS_DIR", kind: "path" },
  { flag: "--relations", env: "FOLIO_SCENARIO_RELATIONS", kind: "value" },
  { flag: "--relations-depth", env: "FOLIO_SCENARIO_RELATIONS_DEPTH", kind: "value" },
  { flag: "--feature-weights", env: "FOLIO_SCENARIO_FEATURE_WEIGHTS", kind: "path" },
  { flag: "--long-fuzz", env: "FOLIO_SCENARIO_LONG_FUZZ", kind: "value" },
  { flag: "--long-steps", env: "FOLIO_SCENARIO_LONG_STEPS", kind: "value" },
  { flag: "--long-output", env: "FOLIO_SCENARIO_LONG_OUTPUT", kind: "path" },
  { flag: "--public-corpus-fuzz", env: "FOLIO_SCENARIO_PUBLIC_CORPUS_FUZZ", kind: "value" },
  { flag: "--public-corpus-steps", env: "FOLIO_SCENARIO_PUBLIC_CORPUS_STEPS", kind: "value" },
  { flag: "--public-corpus-output", env: "FOLIO_SCENARIO_PUBLIC_CORPUS_OUTPUT", kind: "path" },
  {
    flag: "--public-corpus-fixture",
    env: "FOLIO_SCENARIO_PUBLIC_CORPUS_FIXTURE",
    kind: "value",
  },
] as const satisfies readonly {
  flag: `--${string}`;
  env: `FOLIO_SCENARIO_${string}`;
  kind: "value" | "path";
}[];

/** Render environment replay inputs as CLI flags. */
export const consumerScenarioReplayFlags = (
  environment: Readonly<Record<string, string | undefined>>,
  quote: (value: string) => string,
): string =>
  CONSUMER_SCENARIO_REPLAY_INPUTS.flatMap(({ env, flag }) => {
    const value = environment[env];
    return value === undefined ? [] : [`${flag} ${quote(value)}`];
  }).join(" ");
