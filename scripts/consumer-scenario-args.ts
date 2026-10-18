import { panic } from "better-result";
import path from "node:path";

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

export const parseConsumerArgs = (argv: readonly string[], repoRoot: string): ConsumerArgs => {
  const args: ConsumerArgs = {
    tarballs: null,
    packOnly: null,
    keep: false,
    typecheck: false,
    only: null,
    coverageOut: path.join(repoRoot, "test-results", "consumer-scenarios-coverage.json"),
    featureCoverageOut: path.join(repoRoot, "test-results", "consumer-scenarios-features.json"),
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
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    const value = (): string => argv[++index] ?? panic(`consumer-scenarios: ${arg} needs a value`);
    if (arg === "--tarballs") args.tarballs = path.resolve(value());
    else if (arg === "--pack-only") args.packOnly = path.resolve(value());
    else if (arg === "--keep") args.keep = true;
    else if (arg === "--typecheck") args.typecheck = true;
    else if (arg === "--only") args.only = value();
    else if (arg === "--coverage-out") args.coverageOut = path.resolve(value());
    else if (arg === "--feature-coverage-out") args.featureCoverageOut = path.resolve(value());
    else if (arg === "--seed") args.seed = value();
    else if (arg === "--steps") args.steps = value();
    else if (arg === "--runs") args.runs = value();
    else if (arg === "--collision-runs") args.collisionRuns = value();
    else if (arg === "--swarm") args.swarm = value();
    else if (arg === "--flow") args.flow = value();
    else if (arg === "--public-corpus-dir") args.publicCorpusDir = path.resolve(value());
    else if (arg === "--relations") args.relations = value();
    else if (arg === "--relations-depth") args.relationsDepth = value();
    else if (arg === "--feature-weights") args.featureWeights = path.resolve(value());
    else if (arg === "--long-fuzz") args.longFuzz = value();
    else if (arg === "--long-steps") args.longSteps = value();
    else if (arg === "--long-output") args.longOutput = path.resolve(value());
    else if (arg === "--public-corpus-fuzz") args.publicCorpusFuzz = value();
    else if (arg === "--public-corpus-steps") args.publicCorpusSteps = value();
    else if (arg === "--public-corpus-output") args.publicCorpusOutput = path.resolve(value());
    else if (arg === "--public-corpus-fixture") args.publicCorpusFixture = value();
    else if (arg === "--") args.files.push(...argv.slice(index + 1));
    else if (arg !== undefined && !arg.startsWith("-")) args.files.push(arg);
    else panic(`consumer-scenarios: unknown argument ${arg}`);
    if (arg === "--") break;
  }
  return args;
};

const INPUT_ENV = {
  seed: "FOLIO_SCENARIO_SEED",
  steps: "FOLIO_SCENARIO_FUZZ_STEPS",
  runs: "FOLIO_SCENARIO_FUZZ_RUNS",
  collisionRuns: "FOLIO_SCENARIO_COLLISION_RUNS",
  swarm: "FOLIO_SCENARIO_SWARM",
  flow: "FOLIO_SCENARIO_FLOW",
  publicCorpusDir: "FOLIO_SCENARIO_PUBLIC_CORPUS_DIR",
  relations: "FOLIO_SCENARIO_RELATIONS",
  relationsDepth: "FOLIO_SCENARIO_RELATIONS_DEPTH",
  featureWeights: "FOLIO_SCENARIO_FEATURE_WEIGHTS",
  longFuzz: "FOLIO_SCENARIO_LONG_FUZZ",
  longSteps: "FOLIO_SCENARIO_LONG_STEPS",
  longOutput: "FOLIO_SCENARIO_LONG_OUTPUT",
  publicCorpusFuzz: "FOLIO_SCENARIO_PUBLIC_CORPUS_FUZZ",
  publicCorpusSteps: "FOLIO_SCENARIO_PUBLIC_CORPUS_STEPS",
  publicCorpusOutput: "FOLIO_SCENARIO_PUBLIC_CORPUS_OUTPUT",
  publicCorpusFixture: "FOLIO_SCENARIO_PUBLIC_CORPUS_FIXTURE",
} as const;

/** Resolve replay inputs with CLI flags taking precedence over CI environment variables. */
export const consumerScenarioInputs = (
  args: ConsumerArgs,
  environment: Readonly<Record<string, string | undefined>>,
) =>
  Object.fromEntries(
    Object.entries(INPUT_ENV).map(([option, name]) => [
      option,
      args[option as keyof typeof INPUT_ENV] ?? environment[name] ?? null,
    ]),
  ) as Record<keyof typeof INPUT_ENV, string | null>;

if (import.meta.main) {
  console.log(JSON.stringify(parseConsumerArgs(process.argv.slice(2), process.cwd())));
}
