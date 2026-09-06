/**
 * The compare benchmark.
 *
 * ```sh
 * bun benchmarks/compare/run.ts                       # generated corpus
 * bun benchmarks/compare/run.ts --sizes s,m --quick   # a fast loop while editing
 * bun benchmarks/compare/run.ts --corpus path/to/dir  # plus an external corpus
 * bun benchmarks/compare/run.ts --baseline            # record product digests
 * bun benchmarks/compare/run.ts --check               # prove nothing changed
 * ```
 *
 * Every configuration runs in its own process, so a large document class
 * cannot hand the next one a warmed JIT or a grown heap. The parent enumerates
 * configurations, spawns a child per configuration, and aggregates the JSON
 * each child prints.
 */

import { spawnSync } from "node:child_process";
import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { loadavg } from "node:os";
import path from "node:path";

import { compareDocx } from "@stll/folio-core/compare/compare";

import { loadCorpus, readDocument, type CorpusPair } from "./corpus";
import {
  buildDocumentPackage,
  DOCUMENT_CLASSES,
  DOCUMENT_SIZES,
  type DocumentClass,
  type DocumentSize,
} from "./documents";
import { checkInvariants, type InvariantOutcome } from "./invariants";
import { measureHeapGrowth, sample, summarize, type Distribution } from "./measure";
import { zipPackage } from "./package-xml";
import {
  classifyRefusal,
  REFUSAL_BUCKETS,
  summarizeRefusals,
  type PairOutcome,
} from "./refusals";
import { COMPARE_STAGES, runStagedCompare, type CompareStage } from "./stages";
import { applyVariant, EDIT_VARIANTS, type EditVariant } from "./variants";
import { PACKAGE_VALIDATOR_HINT, resolvePackageValidator } from "./validator";

const OPTIONS = { author: "folio compare benchmark", timestamp: "2000-01-01T00:00:00.000Z" };

const DIGESTS_FILE = path.join(import.meta.dir, "digests.json");

type Configuration = { documentClass: DocumentClass; size: DocumentSize; variant: EditVariant };

const configurationId = ({ documentClass, size, variant }: Configuration): string =>
  `${documentClass}/${size}/${variant}`;

const parseConfigurationId = (id: string): Configuration => {
  const [documentClass, size, variant] = id.split("/");
  const configuration = DOCUMENT_CLASSES.find((name) => name === documentClass);
  const sizeName = DOCUMENT_SIZES.find((name) => name === size);
  const variantName = EDIT_VARIANTS.find((name) => name === variant);
  if (!configuration || !sizeName || !variantName) {
    throw new Error(`Not a configuration id: ${id}`);
  }
  return { documentClass: configuration, size: sizeName, variant: variantName };
};

type MeasuredCase = {
  id: string;
  status: "measured";
  baseBytes: number;
  baseBlocks: number;
  changes: number;
  wall: Distribution;
  stages: Record<CompareStage, Distribution>;
  heapGrowthBytes: number;
  /** Distinct reasons the comparison gave for parts it did not look at. */
  unsupported: readonly string[];
  invariants: readonly InvariantOutcome[];
  digests: { buffer: string; changes: string };
  /** Set only for external-corpus cases that publish their own count. */
  expectedChanges: number | null;
};

type SkippedCase = { id: string; status: "skipped"; detail: string };
type FailedCase = { id: string; status: "failed"; detail: string };

type CaseResult = MeasuredCase | SkippedCase | FailedCase;

type Report = {
  machine: { platform: string; arch: string; cpus: number; bun: string };
  /**
   * One-minute host load before and after the run. A run taken while the
   * machine was busy with something else is not comparable with one taken
   * idle, and the only honest way to know is to record it.
   */
  loadAverage: { before: number; after: number };
  warmups: number;
  iterations: number;
  cases: readonly CaseResult[];
};

type PairBytes = { base: ArrayBuffer; target: ArrayBuffer };

const buildGeneratedPair = async ({
  documentClass,
  size,
  variant,
}: Configuration): Promise<PairBytes | null> => {
  const parts = buildDocumentPackage({ documentClass, size });
  const targetParts = applyVariant({ parts, variant });
  if (targetParts === null) {
    return null;
  }
  return { base: await zipPackage(parts), target: await zipPackage(targetParts) };
};

type MeasurePairOptions = {
  id: string;
  pair: PairBytes;
  warmups: number;
  iterations: number;
  expectedChanges: number | null;
  expectation: "identical" | "different";
};

const measurePair = async ({
  id,
  pair,
  warmups,
  iterations,
  expectedChanges,
  expectation,
}: MeasurePairOptions): Promise<CaseResult> => {
  const run = async () => await runStagedCompare(pair.base, pair.target, OPTIONS);
  const probe = await run();
  if (probe.status === "failed") {
    return { id, status: "failed", detail: `${probe.stage}: ${probe.error}` };
  }

  const { wall, last } = await sample({ run, warmups, iterations });
  if (last.status === "failed") {
    return { id, status: "failed", detail: `${last.stage}: ${last.error}` };
  }

  const stageSamples: Record<CompareStage, number[]> = {
    parse: [],
    align: [],
    apply: [],
    serialize: [],
  };
  for (let index = 0; index < iterations; index++) {
    const measured = await run();
    if (measured.status === "failed") {
      return { id, status: "failed", detail: `${measured.stage}: ${measured.error}` };
    }
    for (const stage of COMPARE_STAGES) {
      stageSamples[stage].push(measured.durations[stage]);
    }
  }

  const heapGrowthBytes = await measureHeapGrowth(run);
  const unsupported = [...new Set(last.unsupported.map(({ reason }) => reason))].toSorted();
  const { outcomes, digests } = await checkInvariants({
    base: pair.base,
    target: pair.target,
    redlined: last.buffer,
    changes: last.changes,
    unsupported,
    expectation,
    options: OPTIONS,
    validate: resolvePackageValidator(),
  });

  return {
    id,
    status: "measured",
    baseBytes: pair.base.byteLength,
    baseBlocks: last.baseBlocks,
    changes: last.changes.length,
    wall: summarize(wall),
    stages: {
      parse: summarize(stageSamples.parse),
      align: summarize(stageSamples.align),
      apply: summarize(stageSamples.apply),
      serialize: summarize(stageSamples.serialize),
    },
    heapGrowthBytes,
    unsupported,
    invariants: outcomes,
    digests,
    expectedChanges,
  };
};

const corpusPairBytes = (pair: CorpusPair): PairBytes => ({
  base: readDocument(pair.basePath),
  target: readDocument(pair.targetPath),
});

/**
 * Compare every pair of an external corpus once and report what the refusals
 * were refused for.
 *
 * Unlike the measurement modes this runs in one process: a refusal is a yes or
 * no that no warm JIT can change, and a process per pair would turn a
 * half-minute pass over the whole corpus into ten minutes of spawning.
 */
const runRefusals = async (options: CliOptions): Promise<number> => {
  if (options.corpusDirectory === null) {
    console.log("--refusals needs --corpus <dir>; the repository ships no corpus of its own.");
    return 1;
  }
  const corpus = loadCorpus(options.corpusDirectory);
  if (corpus.status === "empty") {
    console.log(`corpus skipped: ${corpus.detail}`);
    return 1;
  }

  const outcomes: PairOutcome[] = [];
  for (const pair of corpus.pairs) {
    if (options.filter !== null && !pair.id.includes(options.filter)) {
      continue;
    }
    const { base, target } = corpusPairBytes(pair);
    const result = await compareDocx(base, target, OPTIONS);
    outcomes.push(
      result.isOk()
        ? {
            id: pair.id,
            status: "produced",
            changes: result.value.changes.length,
            unsupported: [...new Set(result.value.unsupported.map(({ reason }) => reason))],
          }
        : { id: pair.id, status: "refused", ...classifyRefusal(result.error) },
    );
  }

  const refused = outcomes.filter(({ status }) => status === "refused").length;
  const share = outcomes.length === 0 ? 0 : (refused / outcomes.length) * 100;
  console.log(
    `\n${String(outcomes.length)} pairs, ${String(outcomes.length - refused)} produced, ` +
      `${String(refused)} refused (${share.toFixed(1)}%).\n`,
  );
  console.log("| Bucket | Documents | What it is | A representative shape |");
  console.log("| ------ | --------: | ---------- | ---------------------- |");
  for (const { bucket, count, shape } of summarizeRefusals(outcomes)) {
    console.log(`| \`${bucket}\` | ${String(count)} | ${REFUSAL_BUCKETS[bucket]} | ${shape} |`);
  }

  if (options.out !== null) {
    writeFileSync(options.out, `${JSON.stringify(outcomes, null, 2)}\n`);
    console.log(`\nWrote ${options.out}`);
  }
  return 0;
};

type CliOptions = {
  configuration: string | null;
  corpusPair: string | null;
  corpusDirectory: string | null;
  filter: string | null;
  sizes: readonly DocumentSize[];
  warmups: number;
  iterations: number;
  mode: "measure" | "baseline" | "check" | "refusals";
  out: string | null;
};

const modeOf = (argv: readonly string[]): CliOptions["mode"] => {
  if (argv.includes("--baseline")) {
    return "baseline";
  }
  if (argv.includes("--refusals")) {
    return "refusals";
  }
  return argv.includes("--check") ? "check" : "measure";
};

const parseArguments = (argv: readonly string[]): CliOptions => {
  const read = (name: string): string | null => {
    const index = argv.indexOf(`--${name}`);
    return index === -1 ? null : (argv[index + 1] ?? null);
  };
  const quick = argv.includes("--quick");
  const sizes = read("sizes")
    ?.split(",")
    .flatMap((value) => DOCUMENT_SIZES.filter((size) => size === value));
  return {
    configuration: read("configuration"),
    corpusPair: read("corpus-pair"),
    corpusDirectory: read("corpus"),
    filter: read("filter"),
    sizes: sizes && sizes.length > 0 ? sizes : DOCUMENT_SIZES,
    warmups: Number(read("warmups") ?? (quick ? 1 : 4)),
    iterations: Number(read("iterations") ?? (quick ? 3 : 9)),
    mode: modeOf(argv),
    out: read("out"),
  };
};

const runChild = async (options: CliOptions): Promise<void> => {
  const shared = {
    warmups: options.warmups,
    iterations: options.iterations,
  };
  if (options.corpusPair !== null && options.corpusDirectory !== null) {
    const corpus = loadCorpus(options.corpusDirectory);
    const pair =
      corpus.status === "loaded"
        ? corpus.pairs.find(({ id }) => id === options.corpusPair)
        : undefined;
    if (!pair) {
      process.stdout.write(
        `${JSON.stringify({ id: options.corpusPair, status: "skipped", detail: "pair not found" })}\n`,
      );
      return;
    }
    const result = await measurePair({
      id: pair.id,
      pair: corpusPairBytes(pair),
      expectedChanges: pair.expectedChanges,
      // An external corpus pairs two documents someone else chose; the harness
      // cannot claim they differ.
      expectation: "identical",
      ...shared,
    });
    process.stdout.write(`${JSON.stringify(result)}\n`);
    return;
  }

  const id = options.configuration ?? "";
  const configuration = parseConfigurationId(id);
  const pair = await buildGeneratedPair(configuration);
  if (pair === null) {
    process.stdout.write(
      `${JSON.stringify({ id, status: "skipped", detail: "the variant does not apply to this class" })}\n`,
    );
    return;
  }
  const result = await measurePair({
    id,
    pair,
    expectedChanges: null,
    expectation: configuration.variant === "identical" ? "identical" : "different",
    ...shared,
  });
  process.stdout.write(`${JSON.stringify(result)}\n`);
};

type ChildRequest = { id: string; argv: readonly string[] };

const runInChildProcess = (request: ChildRequest, options: CliOptions): CaseResult => {
  const result = spawnSync(
    process.execPath,
    [
      import.meta.filename,
      ...request.argv,
      "--warmups",
      String(options.warmups),
      "--iterations",
      String(options.iterations),
    ],
    { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 },
  );
  if (result.error || result.status !== 0) {
    return {
      id: request.id,
      status: "failed",
      detail: (result.stderr || result.error?.message || "child exited non-zero")
        .trim()
        .slice(0, 400),
    };
  }
  const line = result.stdout.trim().split("\n").at(-1) ?? "";
  return JSON.parse(line) as CaseResult;
};

const milliseconds = (value: number): string => `${value.toFixed(1)}ms`;

const printCase = (result: CaseResult): void => {
  if (result.status !== "measured") {
    console.log(`${result.id.padEnd(34)} ${result.status}: ${result.detail}`);
    return;
  }
  const stages = COMPARE_STAGES.map(
    (stage) => `${stage.slice(0, 3)} ${milliseconds(result.stages[stage].median)}`,
  ).join("  ");
  const failed = result.invariants.filter(({ status }) => status === "failed");
  const flag = failed.length === 0 ? "ok" : `BROKEN(${failed.map((o) => o.invariant).join(",")})`;
  console.log(
    `${result.id.padEnd(34)} ${String(result.baseBlocks).padStart(5)} blocks ` +
      `${milliseconds(result.wall.median).padStart(9)} [${stages}] ` +
      `${String(result.changes).padStart(5)} changes ` +
      `${(result.heapGrowthBytes / 1024 / 1024).toFixed(1).padStart(6)}MB ${flag}`,
  );
};

type Digests = Record<string, { buffer: string; changes: string }>;

const readDigests = (): Digests =>
  existsSync(DIGESTS_FILE) ? (JSON.parse(readFileSync(DIGESTS_FILE, "utf8")) as Digests) : {};

const applyDigestMode = (report: Report, mode: CliOptions["mode"]): number => {
  const measured = report.cases.filter(
    (result): result is MeasuredCase => result.status === "measured",
  );
  if (mode === "baseline") {
    const digests: Digests = {};
    for (const { id, digests: value } of measured.toSorted((left, right) =>
      left.id.localeCompare(right.id),
    )) {
      digests[id] = value;
    }
    writeFileSync(DIGESTS_FILE, `${JSON.stringify(digests, null, 2)}\n`);
    console.log(`\nWrote ${String(measured.length)} digests to ${DIGESTS_FILE}`);
    return 0;
  }
  if (mode !== "check") {
    return 0;
  }
  const recorded = readDigests();
  const drifted = measured.filter(({ id, digests }) => {
    const previous = recorded[id];
    return (
      previous !== undefined &&
      (previous.buffer !== digests.buffer || previous.changes !== digests.changes)
    );
  });
  const missing = measured.filter(({ id }) => recorded[id] === undefined);
  for (const { id } of drifted) {
    console.log(`digest changed: ${id}`);
  }
  for (const { id } of missing) {
    console.log(`no recorded digest: ${id}`);
  }
  console.log(
    drifted.length === 0
      ? `\nEvery product matches the recorded digest (${String(measured.length - missing.length)} checked).`
      : `\n${String(drifted.length)} products differ from the recorded digests.`,
  );
  return drifted.length === 0 ? 0 : 1;
};

const runParent = (options: CliOptions): number => {
  const loadBefore = loadavg().at(0) ?? Number.NaN;
  const requests: ChildRequest[] = [];
  for (const documentClass of DOCUMENT_CLASSES) {
    for (const size of options.sizes) {
      for (const variant of EDIT_VARIANTS) {
        const id = configurationId({ documentClass, size, variant });
        if (options.filter !== null && !id.includes(options.filter)) {
          continue;
        }
        requests.push({ id, argv: ["--configuration", id] });
      }
    }
  }

  if (options.corpusDirectory !== null) {
    const corpus = loadCorpus(options.corpusDirectory);
    if (corpus.status === "empty") {
      console.log(`corpus skipped: ${corpus.detail}`);
    } else {
      for (const { id } of corpus.pairs) {
        if (options.filter !== null && !id.includes(options.filter)) {
          continue;
        }
        requests.push({
          id,
          argv: ["--corpus", options.corpusDirectory, "--corpus-pair", id],
        });
      }
    }
  }

  if (resolvePackageValidator() === null) {
    console.log(`schema validation skipped; build it with: ${PACKAGE_VALIDATOR_HINT}\n`);
  }

  const cases: CaseResult[] = [];
  for (const request of requests) {
    const result = runInChildProcess(request, options);
    printCase(result);
    cases.push(result);
  }

  const report: Report = {
    machine: {
      platform: process.platform,
      arch: process.arch,
      cpus: navigator.hardwareConcurrency,
      bun: Bun.version,
    },
    loadAverage: { before: loadBefore, after: loadavg().at(0) ?? Number.NaN },
    warmups: options.warmups,
    iterations: options.iterations,
    cases,
  };
  if (options.out !== null) {
    writeFileSync(options.out, `${JSON.stringify(report, null, 2)}\n`);
    console.log(`\nWrote ${options.out}`);
  }

  const broken = cases.filter(
    (result) =>
      result.status === "failed" ||
      (result.status === "measured" && result.invariants.some(({ status }) => status === "failed")),
  );
  console.log(
    `\n${String(cases.length)} cases, ${String(broken.length)} with a failing invariant or error.` +
      ` Host load ${report.loadAverage.before.toFixed(1)} -> ${report.loadAverage.after.toFixed(1)}.`,
  );
  const digestStatus = applyDigestMode(report, options.mode);
  return broken.length === 0 ? digestStatus : 1;
};

const options = parseArguments(process.argv.slice(2));
if (options.configuration !== null || options.corpusPair !== null) {
  await runChild(options);
} else if (options.mode === "refusals") {
  process.exitCode = await runRefusals(options);
} else {
  process.exitCode = runParent(options);
}
