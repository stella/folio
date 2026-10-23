import { panic } from "better-result";
import path from "node:path";

import { CONSUMER_SCENARIO_REPLAY_INPUTS } from "../test/consumer-scenarios/support/replay-inputs";

export {
  CONSUMER_SCENARIO_REPLAY_INPUTS,
  consumerScenarioReplay,
  consumerScenarioReplayInputs,
} from "../test/consumer-scenarios/support/replay-inputs";

const optionForFlag = (flag: `--${string}`): string =>
  flag
    .slice(2)
    .replaceAll(/-([a-z])/gu, (_, letter: string) => letter.toUpperCase());

export type ConsumerArgs = {
  tarballs: string | null;
  packOnly: string | null;
  keep: boolean;
  typecheck: boolean;
  only: string | null;
  coverageOut: string;
  featureCoverageOut: string;
  files: string[];
  seed: string | null;
  steps: string | null;
  runs: string | null;
  collisionRuns: string | null;
  swarm: string | null;
  flow: string | null;
  publicCorpusDir: string | null;
  relations: string | null;
  relationsDepth: string | null;
  featureWeights: string | null;
  longFuzz: string | null;
  longSteps: string | null;
  longOutput: string | null;
  publicCorpusFuzz: string | null;
  publicCorpusSteps: string | null;
  publicCorpusOutput: string | null;
  publicCorpusFixture: string | null;
};

export const parseConsumerArgs = (
  argv: readonly string[],
  repoRoot: string,
): ConsumerArgs => {
  const args: ConsumerArgs = {
    tarballs: null,
    packOnly: null,
    keep: false,
    typecheck: false,
    only: null,
    coverageOut: path.join(
      repoRoot,
      "test-results",
      "consumer-scenarios-coverage.json",
    ),
    featureCoverageOut: path.join(
      repoRoot,
      "test-results",
      "consumer-scenarios-features.json",
    ),
    files: [],
    seed: null,
    steps: null,
    runs: null,
    collisionRuns: null,
    swarm: null,
    flow: null,
    publicCorpusDir: null,
    relations: null,
    relationsDepth: null,
    featureWeights: null,
    longFuzz: null,
    longSteps: null,
    longOutput: null,
    publicCorpusFuzz: null,
    publicCorpusSteps: null,
    publicCorpusOutput: null,
    publicCorpusFixture: null,
  };
  const seenReplayFlags = new Set<string>();
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const value = (): string =>
      argv[++index] ?? panic(`consumer-scenarios: ${arg} needs a value`);
    if (arg === "--tarballs") args.tarballs = path.resolve(value());
    else if (arg === "--pack-only") args.packOnly = path.resolve(value());
    else if (arg === "--keep") args.keep = true;
    else if (arg === "--typecheck") args.typecheck = true;
    else if (arg === "--only") args.only = value();
    else if (arg === "--coverage-out") args.coverageOut = path.resolve(value());
    else if (arg === "--feature-coverage-out")
      args.featureCoverageOut = path.resolve(value());
    else if (arg !== undefined) {
      const input = CONSUMER_SCENARIO_REPLAY_INPUTS.find(
        ({ flag }) => flag === arg || arg.startsWith(`${flag}=`),
      );
      if (input !== undefined) {
        if (seenReplayFlags.has(input.flag)) {
          panic(`consumer-scenarios: repeated replay flag ${input.flag}`);
        }
        seenReplayFlags.add(input.flag);
        const inputValue =
          arg === input.flag ? value() : arg.slice(input.flag.length + 1);
        Object.assign(args, {
          [optionForFlag(input.flag)]:
            input.kind === "path" ? path.resolve(inputValue) : inputValue,
        });
      } else if (arg === "--") args.files.push(...argv.slice(index + 1));
      else if (!arg.startsWith("-")) args.files.push(arg);
      else panic(`consumer-scenarios: unknown argument ${arg}`);
    }
    if (arg === "--") break;
  }
  return args;
};

/** Resolve replay inputs with CLI flags taking precedence over CI environment variables. */
export const consumerScenarioInputs = (
  args: ConsumerArgs,
  environment: Readonly<Record<string, string | undefined>>,
) =>
  Object.fromEntries(
    CONSUMER_SCENARIO_REPLAY_INPUTS.map(({ flag, env }) => [
      optionForFlag(flag),
      args[optionForFlag(flag) as keyof ConsumerArgs] ??
        environment[env] ??
        null,
    ]),
  );

if (import.meta.main) {
  console.log(
    JSON.stringify(parseConsumerArgs(process.argv.slice(2), process.cwd())),
  );
}
