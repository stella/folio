/**
 * Attribute the difference between two corpus runs, per family and per file.
 *
 * Both runs must have been produced with `corpus-gate.ts run --per-file`: a
 * census without the per-file rows carries three examples per signature, which
 * is not enough to say which files a row gained.
 *
 * Usage:
 *   bun scripts/corpus-diff.ts --before a1.json [--before a2.json ...] \
 *                              --after  b1.json [--after  b2.json ...]
 */

import { TaggedError } from "better-result";

import { diffFileSignatures, renderFamilyDeltas } from "./lib/corpus-diff";
import type { CorpusFileSignatures } from "./lib/corpus-file-signatures";

class CorpusDiffError extends TaggedError("CorpusDiffError")<{ message: string }> {}

const flagValues = (args: readonly string[], flag: string): string[] => {
  const values: string[] = [];
  for (const [index, arg] of args.entries()) {
    if (arg !== flag) {
      continue;
    }
    const value = args.at(index + 1);
    if (value === undefined || value.startsWith("--")) {
      throw new CorpusDiffError({ message: `${flag} needs a value` });
    }
    values.push(value);
  }
  return values;
};

type CensusWithFileSignatures = { fileSignatures?: CorpusFileSignatures[] };

/**
 * A side's rows, refusing a census that has none.
 *
 * Reading a `--per-file`-less census as an empty side would report every
 * signature as fixed or introduced, which is a wrong answer rather than a
 * missing one.
 */
const loadSide = async (
  flag: string,
  paths: readonly string[],
): Promise<CorpusFileSignatures[]> => {
  if (paths.length === 0) {
    throw new CorpusDiffError({ message: `${flag} needs at least one census file` });
  }
  const rows: CorpusFileSignatures[] = [];
  for (const file of paths) {
    // oxlint-disable-next-line no-await-in-loop -- one file at a time keeps peak memory at one census
    const census = (await Bun.file(file).json()) as CensusWithFileSignatures;
    if (census.fileSignatures === undefined) {
      throw new CorpusDiffError({
        message: `${file} carries no per-file rows; rerun that side with \`corpus-gate.ts run --per-file\``,
      });
    }
    rows.push(...census.fileSignatures);
  }
  return rows;
};

const main = async (args: readonly string[]): Promise<void> => {
  const [before, after] = await Promise.all([
    loadSide("--before", flagValues(args, "--before")),
    loadSide("--after", flagValues(args, "--after")),
  ]);
  process.stdout.write(`${renderFamilyDeltas(diffFileSignatures(before, after))}\n`);
};

if (import.meta.main) {
  main(process.argv.slice(2)).catch((cause: unknown) => {
    process.stderr.write(`${cause instanceof Error ? cause.message : String(cause)}\n`);
    process.exitCode = 1;
  });
}
