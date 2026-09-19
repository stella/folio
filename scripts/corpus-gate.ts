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
 *   bun scripts/corpus-gate.ts run [--shard k/n] [--concurrency N] [--timeout MS]
 *                                  [--tiers 1,2] [--invariant-budget MS] [--file-budget MS]
 *                                  [--only ID[,ID...]] [--out FILE] [--check]
 *   bun scripts/corpus-gate.ts check <census.json...>
 *   bun scripts/corpus-gate.ts write-baseline <census.json...>
 *   bun scripts/corpus-gate.ts report <census.json...>
 */

import { mkdir } from "node:fs/promises";
import path from "node:path";

import { TaggedError } from "better-result";

import {
  type CorpusBaseline,
  baselineFromCensus,
  compareToBaseline,
  describeDegradedRun,
  isDegradedRun,
  isFailingViolation,
  renderViolations,
} from "./lib/corpus-baseline";
import {
  CensusBuilder,
  type CorpusCensus,
  type CorpusFileResult,
  evidenceOf,
  mergeCensuses,
  renderCensus,
} from "./lib/corpus-census";
import {
  FAMILY_BASELINE_FAMILIES,
  compareFamilyToBaseline,
  loadFamilyBaseline,
  writeFamilyBaselines,
} from "./lib/corpus-family-baseline";
import {
  FamilyCensusBuilder,
  type FamilyCensus,
  censusWithLateFailures,
  mergeFamilyCensuses,
  renderFamilyCensus,
} from "./lib/corpus-family-census";
import { performanceFailures, medianMsPerMegabyte } from "./lib/corpus-invariants/performance";
import {
  BASELINE_PATH,
  EXPECTED_REFUSALS_PATH,
  type CorpusLock,
  type CorpusTier,
  corpusCacheRoot,
  loadCorpusLock,
  writeJsonFile,
} from "./lib/corpus-manifest";
import {
  type CorpusBudgets,
  type CorpusTask,
  type CorpusTaskOutcome,
  runCorpusPool,
} from "./lib/corpus-pool";
import {
  assertReportOnlyFilesAreLive,
  loadReportOnlyFiles,
  reportOnlyFileIds,
  reportOnlyFilesDigest,
} from "./lib/corpus-report-only";
import { corpusFileId, parseOnlySelection, selectOnly } from "./lib/corpus-selection";
import {
  type ExpectedRefusals,
  compareToExpectedRefusals,
  partitionExpectedRefusals,
  refreshedExpectedRefusals,
  renderExpectedRefusals,
} from "./lib/corpus-refusals";
import {
  describeTiers,
  parseTierSelection,
  selectTiers,
  tierScopedLockDigest,
} from "./lib/corpus-tiers";

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
const REPORTED_SLOW_FILES = 20;
const PROGRESS_INTERVAL = 250;
/** Milliseconds one extended invariant may spend on one file before the overrun is a finding. */
const DEFAULT_INVARIANT_BUDGET_MS = 30_000;
/** Milliseconds all of them together may spend before the rest are skipped. */
const DEFAULT_FILE_BUDGET_MS = 120_000;
/** The stage a file that never answered stopped at, as the census names it. */
const WATCHDOG_STAGE = "the worker deadline";

/**
 * Decide the performance family once the whole run is in.
 *
 * An outlier is a file whose parse costs many times what the corpus costs per
 * megabyte, which is not knowable from the file alone.
 */
const withPerformance = (census: FamilyCensus): FamilyCensus => {
  const median = medianMsPerMegabyte(census.costs);
  return censusWithLateFailures(
    census,
    census.costs.map((cost) => ({
      file: cost.file,
      producer: cost.producer,
      failures: performanceFailures(cost, median),
    })),
  );
};

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

const flagValues = (args: readonly string[], flag: string): string[] => {
  const values: string[] = [];
  for (const [index, arg] of args.entries()) {
    if (arg !== flag) {
      continue;
    }
    const value = args.at(index + 1);
    if (value === undefined || value.startsWith("--")) {
      throw new CorpusGateError({ message: `${flag} needs a value` });
    }
    values.push(value);
  }
  return values;
};

const flagValue = (args: readonly string[], flag: string): string | undefined =>
  flagValues(args, flag).at(0);

type CorpusFileEntry = CorpusTask & { duplicateOf: string | null };

/**
 * Every locked file, in a fixed order, with duplicate content marked.
 *
 * Public suites copy fixtures between projects, so the corpus carries the same
 * bytes under several names. Sharding assigns work by position in this list, so
 * a shard sees the same files however many shards there are, and a duplicate is
 * counted by whichever shard owns it rather than run again.
 */
const buildFileList = async (
  tiers: readonly CorpusTier[],
): Promise<{ entries: CorpusFileEntry[]; lockDigest: string; lock: CorpusLock }> => {
  const full = await loadCorpusLock();
  const lock = selectTiers(full, tiers);
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
  return { entries, lockDigest: tierScopedLockDigest(full, tiers), lock: full };
};

/**
 * Re-decide which of a selection are duplicates, within the selection alone.
 *
 * `buildFileList` marks a duplicate against the whole corpus, which is right
 * for a sharded run and wrong for `--only`: naming one file that happens to
 * share its bytes with another source's copy would otherwise run nothing.
 */
const withinSelectionDuplicates = (entries: readonly CorpusFileEntry[]): CorpusFileEntry[] => {
  const firstIdBySha = new Map<string, string>();
  return entries.map((entry) => {
    const seen = firstIdBySha.get(entry.sha256);
    if (seen === undefined) {
      firstIdBySha.set(entry.sha256, corpusFileId(entry));
      return { ...entry, duplicateOf: null };
    }
    return { ...entry, duplicateOf: seen };
  });
};

type RunOptions = {
  shard: Shard;
  concurrency: number;
  timeoutMs: number;
  outPath: string;
  tiers: readonly CorpusTier[];
  budgets: CorpusBudgets;
  /** `undefined` is the whole corpus; a list narrows the run to those files. */
  only: readonly string[] | undefined;
};

/**
 * One census file carries both censuses.
 *
 * A run produces them from the same observations, and splitting them across two
 * files would let a merge pair a core census with a family census from another
 * shard.
 */
type CorpusCensusFile = CorpusCensus & { family: FamilyCensus };

/** How a file's run ended, before the committed list has its say. */
const resultOfOutcome = (outcome: CorpusTaskOutcome): CorpusFileResult => {
  switch (outcome.kind) {
    case "not-a-docx": {
      return { kind: "not-a-docx", reason: outcome.reason };
    }
    // The worker died on this file: a fact about the file, which gates.
    case "aborted": {
      return { kind: "complete", failures: outcome.failures };
    }
    // The deadline expired, so this file's evidence is as load-dependent as a
    // budget truncation's, and it is treated as one.
    case "watchdog-expired": {
      return { kind: "truncated", failures: outcome.failures, stage: WATCHDOG_STAGE };
    }
    case "checked": {
      return outcome.truncatedAt === undefined
        ? { kind: "complete", failures: outcome.failures }
        : { kind: "truncated", failures: outcome.failures, stage: outcome.truncatedAt };
    }
    default: {
      throw new CorpusGateError({
        message: `unhandled corpus task outcome: ${JSON.stringify(outcome)}`,
      });
    }
  }
};

/**
 * What one file's run means for the ratchet, decided in one place.
 *
 * Membership of the committed list is the only thing that may excuse a file
 * from contributing gating evidence, and it outranks how the run went: a
 * listed file is report-only whether it finished or stopped at a budget, and
 * an unlisted file that stopped is truncated, which degrades the run.
 */
const corpusFileResult = (outcome: CorpusTaskOutcome, listed: boolean): CorpusFileResult => {
  const result = resultOfOutcome(outcome);
  if (!listed || result.kind === "not-a-docx") {
    return result;
  }
  return { kind: "report-only", failures: result.failures };
};

const runGate = async ({
  shard,
  concurrency,
  timeoutMs,
  outPath,
  tiers,
  budgets,
  only,
}: RunOptions): Promise<CorpusCensusFile> => {
  const { entries, lockDigest, lock } = await buildFileList(tiers);
  const reportOnly = await loadReportOnlyFiles();
  // Against the whole lock, not the tier selection: an entry for a file this
  // run does not reach is still an entry that must name a file that exists.
  assertReportOnlyFilesAreLive(reportOnly, lock);
  const reportOnlyIds = reportOnlyFileIds(reportOnly);
  const reportOnlyDigest = reportOnlyFilesDigest(reportOnly);
  const mine =
    only === undefined
      ? entries.filter((_, index) => index % shard.total === shard.index - 1)
      : withinSelectionDuplicates(selectOnly({ entries, patterns: only, idOf: corpusFileId }));
  if (mine.length === 0) {
    throw new CorpusGateError({
      message: "The corpus cache is empty. Run `bun run corpus:fetch` first.",
    });
  }

  const census = new CensusBuilder(lockDigest, reportOnlyDigest);
  const family = new FamilyCensusBuilder(lockDigest, reportOnlyDigest);
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
    budgets,
    onOutcome: (task, outcome) => {
      done += 1;
      if (done % PROGRESS_INTERVAL === 0) {
        process.stderr.write(`  ${done}/${tasks.length} files\n`);
      }
      const file = { sourceId: task.sourceId, path: task.relativePath, sha256: task.sha256 };
      const result = corpusFileResult(outcome, reportOnlyIds.has(corpusFileId(task)));
      census.add(file, result);
      if (outcome.kind === "checked") {
        family.add({
          file,
          bytes: outcome.cost.bytes,
          parseMs: outcome.cost.parseMs,
          peakRssBytes: outcome.cost.peakRssBytes,
          producer: outcome.producer,
          failures: outcome.failures,
          timings: outcome.timings,
          evidence: evidenceOf(result),
        });
      }
    },
  });

  const built = { ...census.build(), family: family.build() };
  await mkdir(path.dirname(outPath), { recursive: true });
  await writeJsonFile(outPath, built);
  const seconds = ((Bun.nanoseconds() - started) / 1e9).toFixed(1);
  process.stdout.write(
    `Corpus gate ${only === undefined ? `shard ${shard.index}/${shard.total}` : `--only, ${mine.length} file(s)`} over ${describeTiers(tiers)} in ${seconds}s -> ${outPath}\n` +
      `${renderCensus(built, REPORTED_SIGNATURES)}\n` +
      `${renderFamilyCensus({ census: withPerformance(built.family), signaturesPerFamily: REPORTED_SIGNATURES, slowestPerStage: REPORTED_SLOW_FILES })}\n`,
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

const loadExpectedRefusals = async (): Promise<ExpectedRefusals> => {
  const file = Bun.file(EXPECTED_REFUSALS_PATH);
  if (!(await file.exists())) {
    // Absent, every deliberate refusal would read as a defect and
    // `write-baseline` would bake those signatures into the baseline.
    throw new CorpusGateError({
      message:
        "corpus/expected-refusals.json is missing. It is committed; restore it rather than running without it.",
    });
  }
  return (await file.json()) as ExpectedRefusals;
};

const loadCensuses = async (paths: readonly string[]): Promise<CorpusCensusFile> => {
  if (paths.length === 0) {
    throw new CorpusGateError({ message: "Pass at least one census file" });
  }
  const censuses = await Promise.all(
    paths.map(async (file) => (await Bun.file(file).json()) as CorpusCensusFile),
  );
  return {
    ...mergeCensuses(censuses),
    family: mergeFamilyCensuses(censuses.map((census) => census.family)),
  };
};

/**
 * The committed list must name files the corpus still carries.
 *
 * Checked wherever a baseline is read or written, not only on a run: an entry
 * that no longer resolves is an exemption nobody reviewed, and the gate has to
 * say so rather than carry it.
 */
const assertReportOnlyListIsLive = async (): Promise<void> => {
  assertReportOnlyFilesAreLive(await loadReportOnlyFiles(), await loadCorpusLock());
};

const checkAgainstBaseline = async (census: CorpusCensusFile): Promise<void> => {
  await assertReportOnlyListIsLive();
  const refusals = await loadExpectedRefusals();
  const { defects, refusals: observedRefusals } = partitionExpectedRefusals(census, refusals);
  const violations = [
    ...compareToBaseline(await loadBaseline(), defects),
    ...compareToExpectedRefusals(refusals, observedRefusals),
  ];
  const family = withPerformance(census.family);
  for (const name of FAMILY_BASELINE_FAMILIES) {
    // oxlint-disable-next-line no-await-in-loop -- one small file per family, read in a fixed order
    violations.push(...compareFamilyToBaseline(await loadFamilyBaseline(name), family));
  }
  const rendered = renderExpectedRefusals(refusals, observedRefusals);
  const failing = violations.filter(isFailingViolation);
  const informational = violations.filter((violation) => !isFailingViolation(violation));
  const truncated = renderTruncated(census);
  if (informational.length > 0) {
    process.stdout.write(
      `Corpus gate: kept ${informational.length} baseline entr(ies) this run was not asked to confirm:\n${renderViolations(informational)}\n`,
    );
  }
  if (failing.length === 0) {
    process.stdout.write(
      `Corpus gate: no change against the baselines (${defects.signatures.length} known defects, ${refusals.entries.length} expected refusals, ${census.reportOnly} report-only files)\n${truncated}${rendered}\n`,
    );
    return;
  }
  process.stderr.write(
    `Corpus gate baseline violations:\n${renderViolations(failing)}\n${truncated}`,
  );
  process.exitCode = 1;
};

/** Truncated files get their own heading: they are why a count may be short. */
const renderTruncated = (census: CorpusCensusFile): string =>
  isDegradedRun(census) ? `${describeDegradedRun(census)}\n` : "";

const main = async (args: string[]): Promise<void> => {
  const command = args.at(0);
  const rest = args.slice(1);

  if (command === "run") {
    const shard = parseShard(flagValue(rest, "--shard"));
    const tiers = parseTierSelection(flagValue(rest, "--tiers"));
    const only = parseOnlySelection(flagValues(rest, "--only"));
    // A subset census records neither the signatures the rest of the corpus
    // fires nor the file counts a baseline entry carries, so ratcheting one
    // would report every unselected defect as resolved.
    if (only !== undefined && (rest.includes("--check") || shard.total !== 1)) {
      throw new CorpusGateError({
        message: "--only runs a subset, so it cannot be combined with --check or --shard",
      });
    }
    const outPath =
      flagValue(rest, "--out") ??
      path.join(
        corpusCacheRoot(),
        "reports",
        only === undefined ? `census-${shard.index}-of-${shard.total}.json` : "census-subset.json",
      );
    const census = await runGate({
      shard,
      tiers,
      only,
      budgets: {
        invariantBudgetMs: parsePositiveInteger(
          flagValue(rest, "--invariant-budget"),
          "--invariant-budget",
          DEFAULT_INVARIANT_BUDGET_MS,
        ),
        fileBudgetMs: parsePositiveInteger(
          flagValue(rest, "--file-budget"),
          "--file-budget",
          DEFAULT_FILE_BUDGET_MS,
        ),
      },
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
    await assertReportOnlyListIsLive();
    const census = await loadCensuses(rest);
    // Writing is stricter than comparing: a comparison can tolerate a thin run
    // by keeping what it could not confirm, but a baseline written from one
    // records counts that are low for reasons outside the code, and the next
    // healthy run reads that as a regression.
    if (isDegradedRun(census)) {
      throw new CorpusGateError({
        message: `${describeDegradedRun(census)} A baseline written from this run would undercount.`,
      });
    }
    const refusals = await loadExpectedRefusals();
    const { defects, refusals: observedRefusals } = partitionExpectedRefusals(census, refusals);
    const family = withPerformance(census.family);
    await writeJsonFile(BASELINE_PATH, baselineFromCensus(defects));
    await writeFamilyBaselines(family);
    // Counts only: which signatures are expected refusals, and why, is a
    // decision recorded by hand in corpus/expected-refusals.json.
    await writeJsonFile(
      EXPECTED_REFUSALS_PATH,
      refreshedExpectedRefusals(refusals, observedRefusals),
    );
    process.stdout.write(
      `corpus/baseline.json written: ${defects.signatures.length} defect signatures, ${refusals.entries.length} expected refusals\n` +
        `corpus/baselines/: ${family.signatures.length} signatures across ${FAMILY_BASELINE_FAMILIES.length} families\n`,
    );
    return;
  }

  if (command === "report") {
    const census = await loadCensuses(rest);
    process.stdout.write(
      `${renderCensus(census, REPORTED_SIGNATURES)}\n${renderFamilyCensus({
        census: withPerformance(census.family),
        signaturesPerFamily: REPORTED_SIGNATURES,
        slowestPerStage: REPORTED_SLOW_FILES,
      })}\n`,
    );
    return;
  }

  throw new CorpusGateError({
    message: "Usage: bun scripts/corpus-gate.ts [run|check|write-baseline|report]",
  });
};

if (import.meta.main) {
  main(process.argv.slice(2)).catch((cause: unknown) => {
    process.stderr.write(`${cause instanceof Error ? cause.message : String(cause)}\n`);
    process.exitCode = 1;
  });
}
