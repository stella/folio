const CONSUMER_SCENARIO_EXECUTABLE = "bun scripts/consumer-scenarios.ts";

type ReplayCommandOptions = {
  inputs?: Readonly<Record<string, string | null | undefined>>;
  flags?: readonly string[];
  files?: readonly string[];
};

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

const optionForFlag = (flag: `--${string}`): string =>
  flag.slice(2).replaceAll(/-([a-z])/gu, (_, letter: string) => letter.toUpperCase());

/** Read the replay inputs present in an environment into one mergeable value set. */
export const consumerScenarioReplayInputs = (
  environment: Readonly<Record<string, string | undefined>>,
): Readonly<Record<string, string>> =>
  Object.fromEntries(
    CONSUMER_SCENARIO_REPLAY_INPUTS.flatMap(({ env, flag }) => {
      const value = environment[env];
      return value === undefined ? [] : [[optionForFlag(flag), value]];
    }),
  );

/** Build a replay from one resolved input set, followed by other flags and file selectors. */
export const consumerScenarioReplay = ({
  inputs = {},
  flags = [],
  files = [],
}: ReplayCommandOptions = {}): string =>
  [
    CONSUMER_SCENARIO_EXECUTABLE,
    ...CONSUMER_SCENARIO_REPLAY_INPUTS.flatMap(({ flag }) => {
      const value = inputs[optionForFlag(flag)];
      if (value == null) return [];
      const quotedValue = value.replaceAll("'", `'\\''`);
      return value.startsWith("-") ? [`'${flag}=${quotedValue}'`] : [`${flag} '${quotedValue}'`];
    }),
    ...flags.filter((flag) => flag !== ""),
    ...(files.length === 0 ? [] : ["--", ...files]),
  ].join(" ");
