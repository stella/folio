import { panic } from "better-result";
import path from "node:path";

type Args = {
  tarballs: string | null;
  packOnly: string | null;
  keep: boolean;
  typecheck: boolean;
  only: string | null;
  coverageOut: string;
  featureCoverageOut: string;
  files: string[];
};

export const parseConsumerArgs = (argv: readonly string[], repoRoot: string): Args => {
  const args: Args = {
    tarballs: null,
    packOnly: null,
    keep: false,
    typecheck: false,
    only: null,
    coverageOut: path.join(repoRoot, "test-results", "consumer-scenarios-coverage.json"),
    featureCoverageOut: path.join(repoRoot, "test-results", "consumer-scenarios-features.json"),
    files: [],
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
    else if (arg === "--") args.files.push(...argv.slice(index + 1));
    else if (arg !== undefined && !arg.startsWith("-")) args.files.push(arg);
    else panic(`consumer-scenarios: unknown argument ${arg}`);
    if (arg === "--") break;
  }
  return args;
};

if (import.meta.main) {
  console.log(JSON.stringify(parseConsumerArgs(process.argv.slice(2), process.cwd())));
}
