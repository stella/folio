/**
 * Run folio's entry points over the public DOCX corpus and ratchet the result.
 *
 * The property tests generate the inputs someone thought to describe. This gate
 * runs the ones nobody did: thousands of real packages written by Word and by
 * every other producer, carried by public test suites. It asserts invariants
 * rather than expected output, so a file needs no oracle to be useful, and it
 * groups what fails into signatures so a census of thousands of files reads as
 * a short list of defects.
 *
 * Usage:
 *   bun scripts/corpus-gate.ts run [--shard k/n] [--concurrency N] [--timeout MS] [--out FILE] [--check]
 *   bun scripts/corpus-gate.ts check <census.json...>
 *   bun scripts/corpus-gate.ts write-baseline <census.json...>
 */

import { mkdir } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

import {
  type CorpusBaseline,
  baselineFromCensus,
  compareToBaseline,
  renderViolations,
} from "./lib/corpus-baseline";
import { CensusBuilder, type CorpusCensus, mergeCensuses, renderCensus } from "./lib/corpus-census";
import {
  BASELINE_PATH,
  corpusCacheRoot,
  corpusLockDigest,
  loadCorpusLock,
  writeJsonFile,
} from "./lib/corpus-manifest";
import { type CorpusTask, runCorpusPool } from "./lib/corpus-pool";

class CorpusGateError extends TaggedError("CorpusGateError")<{ message: string }> {}

const DEFAULT_CONCURRENCY = 4;
/**
 * Generous on purpose: the deadline is here to catch a hang, not slowness.
 *
 * The corpus already contains packages that take the better part of a minute on
 * a loaded machine, and a `completes` failure that only reproduces under load
 * would make the ratchet flap instead of reporting a defect.
 */
const DEFAULT_FILE_TIMEOUT_MS = 300_000;
const REPORTED_SIGNATURES = 25;
const PROGRESS_INTERVAL = 250;

type Shard = { index: number; total: number };

const parseShard = (value: string | undefined): Shard => {
  if (value === undefined) {
    return { index: 1, total: 1 };
  }
  const [rawIndex, rawTotal] = value.split("/");
  const index = Number(rawIndex);
  const total = Number(rawTotal);
  if (
    !Number.isInteger(index) ||
    !Number.isInteger(total) ||
    total < 1 ||
    index < 1 ||
    index > total
  ) {
    throw new CorpusGateError({
      message: `--shard expects \`k/n\` with 1 <= k <= n, got \`${value}\``,
    });
  }
  return { index, total };
};

const parsePositiveInteger = (
  value: string | undefined,
  flag: string,
  fallback: number,
): number => {
  if (value === undefined) {
    return fallback;
  }
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1) {
    throw new CorpusGateError({ message: `${flag} expects a positive integer, got \`${value}\`` });
  }
  return parsed;
};

const flagValue = (args: readonly string[], flag: string): string | undefined => {
  const index = args.indexOf(flag);
  if (index === -1) {
    return undefined;
  }
  const value = args.at(index + 1);
  if (value === undefined || value.startsWith("--")) {
    throw new CorpusGateError({ message: `${flag} needs a value` });
  }
  return value;
};

type CorpusFileEntry = CorpusTask & { duplicateOf: string | null };

/**
 * Every locked file, in a fixed order, with duplicate content marked.
 *
 * Public suites copy fixtures between projects, so the corpus carries the same
 * bytes under several names. Sharding assigns work by position in this list, so
 * a shard sees the same files however many shards there are, and a duplicate is
 * counted by whichever shard owns it rather than run again.
 */
const buildFileList = async (): Promise<{ entries: CorpusFileEntry[]; lockDigest: string }> => {
  const lock = await loadCorpusLock();
  const cacheRoot = corpusCacheRoot();
  const entries: CorpusFileEntry[] = [];
  const firstPathBySha = new Map<string, string>();
  for (const source of [...lock.sources].sort((left, right) => (left.id < right.id ? -1 : 1))) {
    for (const file of [...source.files].sort((left, right) => (left.path < right.path ? -1 : 1))) {
      const seen = firstPathBySha.get(file.sha256);
      const identifier = `${source.id}/${file.path}`;
      if (seen === undefined) {
        firstPathBySha.set(file.sha256, identifier);
      }
      entries.push({
        sourceId: source.id,
        relativePath: file.path,
        sha256: file.sha256,
        absolutePath: path.join(cacheRoot, "sources", source.id, file.path),
        duplicateOf: seen ?? null,
      });
    }
  }
  return { entries, lockDigest: corpusLockDigest(lock) };
};

type RunOptions = {
  shard: Shard;
  concurrency: number;
  timeoutMs: number;
  outPath: string;
};

const runGate = async ({
  shard,
  concurrency,
  timeoutMs,
  outPath,
}: RunOptions): Promise<CorpusCensus> => {
  const { entries, lockDigest } = await buildFileList();
  const mine = entries.filter((_, index) => index % shard.total === shard.index - 1);
  if (mine.length === 0) {
    throw new CorpusGateError({
      message: "The corpus cache is empty. Run `bun run corpus:fetch` first.",
    });
  }

  const census = new CensusBuilder(lockDigest);
  const tasks: CorpusTask[] = [];
  for (const entry of mine) {
    if (entry.duplicateOf === null) {
      tasks.push(entry);
      continue;
    }
    census.countDuplicate();
  }

  const started = Bun.nanoseconds();
  let done = 0;
  await runCorpusPool({
    tasks,
    concurrency,
    timeoutMs,
    onOutcome: (task, outcome) => {
      done += 1;
      if (done % PROGRESS_INTERVAL === 0) {
        process.stderr.write(`  ${done}/${tasks.length} files\n`);
      }
      const file = { sourceId: task.sourceId, path: task.relativePath, sha256: task.sha256 };
      if (outcome.kind === "not-a-docx") {
        census.addNotADocx(file, outcome.reason);
        return;
      }
      census.addChecked(file, outcome.failures);
    },
  });

  const built = census.build();
  await mkdir(path.dirname(outPath), { recursive: true });
  await writeJsonFile(outPath, built);
  const seconds = ((Bun.nanoseconds() - started) / 1e9).toFixed(1);
  process.stdout.write(
    `Corpus gate shard ${shard.index}/${shard.total} in ${seconds}s -> ${outPath}\n${renderCensus(built, REPORTED_SIGNATURES)}\n`,
  );
  return built;
};

const loadBaseline = async (): Promise<CorpusBaseline> => {
  const file = Bun.file(BASELINE_PATH);
  if (!(await file.exists())) {
    throw new CorpusGateError({
      message: "corpus/baseline.json is missing. Generate it with `corpus-gate.ts write-baseline`.",
    });
  }
  return (await file.json()) as CorpusBaseline;
};

const loadCensuses = async (paths: readonly string[]): Promise<CorpusCensus> => {
  if (paths.length === 0) {
    throw new CorpusGateError({ message: "Pass at least one census file" });
  }
  const censuses = await Promise.all(
    paths.map(async (file) => (await Bun.file(file).json()) as CorpusCensus),
  );
  return mergeCensuses(censuses);
};

const checkAgainstBaseline = async (census: CorpusCensus): Promise<void> => {
  const violations = compareToBaseline(await loadBaseline(), census);
  if (violations.length === 0) {
    process.stdout.write(
      `Corpus gate: no change against the baseline (${census.failedFiles} known failures)\n`,
    );
    return;
  }
  process.stderr.write(`Corpus gate baseline violations:\n${renderViolations(violations)}\n`);
  process.exitCode = 1;
};

const main = async (args: string[]): Promise<void> => {
  const command = args.at(0);
  const rest = args.slice(1);

  if (command === "run") {
    const shard = parseShard(flagValue(rest, "--shard"));
    const outPath =
      flagValue(rest, "--out") ??
      path.join(corpusCacheRoot(), "reports", `census-${shard.index}-of-${shard.total}.json`);
    const census = await runGate({
      shard,
      concurrency: parsePositiveInteger(
        flagValue(rest, "--concurrency"),
        "--concurrency",
        DEFAULT_CONCURRENCY,
      ),
      timeoutMs: parsePositiveInteger(
        flagValue(rest, "--timeout"),
        "--timeout",
        DEFAULT_FILE_TIMEOUT_MS,
      ),
      outPath,
    });
    if (rest.includes("--check")) {
      await checkAgainstBaseline(census);
    }
    return;
  }

  if (command === "check") {
    await checkAgainstBaseline(await loadCensuses(rest));
    return;
  }

  if (command === "write-baseline") {
    const census = await loadCensuses(rest);
    await writeJsonFile(BASELINE_PATH, baselineFromCensus(census));
    process.stdout.write(
      `corpus/baseline.json written: ${census.signatures.length} signatures over ${census.failedFiles} files\n`,
    );
    return;
  }

  throw new CorpusGateError({
    message: "Usage: bun scripts/corpus-gate.ts [run|check|write-baseline]",
  });
};

if (import.meta.main) {
  main(process.argv.slice(2)).catch((cause: unknown) => {
    process.stderr.write(`${cause instanceof Error ? cause.message : String(cause)}\n`);
    process.exitCode = 1;
  });
}
