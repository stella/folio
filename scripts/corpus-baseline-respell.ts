/**
 * Check a census against the committed baselines as a re-spelling.
 *
 * Naming the carrier in a path segment rewrites every signature the model
 * comparison produces: `content[].content[].preservedAttributes` becomes
 * `content[paragraph].content[run].preservedAttributes`, and the recorded rows
 * match nothing. The only honest way to move the baselines is to re-measure
 * over the same corpus and write what the run observed, which is what
 * `corpus-gate.ts write-baseline` already does.
 *
 * What `write-baseline` cannot say is whether the run that produced the census
 * changed anything besides the spelling. Every row moved, so a defect that
 * appeared and a defect that was fixed both look like ordinary churn in the
 * diff, and the diff is thousands of lines. This reads both sides, erases the
 * carriers from the observed rows, and reports what the re-spelling explains
 * and what it does not:
 *
 *   bun scripts/corpus-baseline-respell.ts census-1.json … census-4.json
 *
 * It writes nothing and fails when a row is unexplained on either side. Bake
 * with `corpus-gate.ts write-baseline` once it passes.
 */

import { TaggedError } from "better-result";

import { type CorpusBaseline, type CorpusBaselineEntry } from "./lib/corpus-baseline";
import { type CensusSignature, type CorpusCensus, mergeCensuses } from "./lib/corpus-census";
import {
  FAMILY_BASELINE_FAMILIES,
  type FamilyBaseline,
  loadFamilyBaseline,
} from "./lib/corpus-family-baseline";
import { type FamilyCensus, mergeFamilyCensuses } from "./lib/corpus-family-census";
import { BASELINE_PATH } from "./lib/corpus-manifest";
import { ELISION, MODEL_PATH_RE } from "./lib/corpus-signature";

class RespellError extends TaggedError("RespellError")<{ message: string }> {}

/** A row from either side, reduced to what a re-spelling comparison needs. */
type Row = { invariant: string; message: string; files: number };

const ARRAY_SEGMENT_RE = /\[[^[\]]*\]/gu;

/**
 * A message with every carrier erased from its model path.
 *
 * Only inside the path: a message can carry brackets of its own, and erasing
 * those would merge defects this is here to tell apart.
 */
const withoutCarriers = (message: string): string => {
  const match = MODEL_PATH_RE.exec(message);
  if (match === null) {
    return message;
  }
  const [path] = match;
  return (
    message.slice(0, match.index) +
    path.replaceAll(ARRAY_SEGMENT_RE, "[]") +
    message.slice(match.index + path.length)
  );
};

/** What a row is compared on: the invariant and its message with the carriers erased. */
const exactKey = ({ invariant, message }: Row): string =>
  `${invariant} | ${withoutCarriers(message)}`;

/**
 * A row reduced to what a comparison reads.
 *
 * `cut` is the part a message cut from the right kept, and `undefined` for
 * every other row. That distinction is the whole of the second comparison: a
 * message cut from the right is a prefix of its full spelling, so a recorded
 * row that lost its tail to the old cap still matches the observed row it grew
 * into. A message that lost its middle instead is a prefix of nothing, so it
 * matches exactly or it is reported.
 */
type Reduced = { row: Row; erased: string; cut: string | undefined };

const reduce = (row: Row): Reduced => {
  const erased = withoutCarriers(row.message);
  return {
    row,
    erased,
    cut: erased.endsWith(ELISION) ? erased.slice(0, -ELISION.length) : undefined,
  };
};

type Side = { label: string };

type RespellReport = {
  exact: number;
  byPrefix: number;
  /** Observed rows no recorded row explains: a defect class that did not exist before. */
  appeared: Row[];
  /** Recorded rows no observed row explains: a class the run no longer produces. */
  vanished: Row[];
};

const sameInvariant = (left: Reduced, right: Reduced): boolean =>
  left.row.invariant === right.row.invariant;

const matchesByPrefix = (row: Reduced, others: readonly Reduced[]): boolean =>
  others.some(
    (other) =>
      sameInvariant(row, other) &&
      ((row.cut !== undefined && other.erased.startsWith(row.cut)) ||
        (other.cut !== undefined && row.erased.startsWith(other.cut))),
  );

const compareSide = (recorded: readonly Row[], observed: readonly Row[]): RespellReport => {
  const recordedRows = recorded.map(reduce);
  const observedRows = observed.map(reduce);
  const recordedExact = new Set(recorded.map(exactKey));
  const observedExact = new Set(observed.map(exactKey));
  const report: RespellReport = { exact: 0, byPrefix: 0, appeared: [], vanished: [] };
  for (const row of observedRows) {
    if (recordedExact.has(exactKey(row.row))) {
      report.exact += 1;
      continue;
    }
    if (matchesByPrefix(row, recordedRows)) {
      report.byPrefix += 1;
      continue;
    }
    report.appeared.push(row.row);
  }
  for (const row of recordedRows) {
    if (!observedExact.has(exactKey(row.row)) && !matchesByPrefix(row, observedRows)) {
      report.vanished.push(row.row);
    }
  }
  return report;
};

const REPORTED_ROWS = 20;

const renderRows = (label: string, rows: readonly Row[]): string =>
  rows.length === 0
    ? ""
    : [
        `  ${rows.length} ${label}:`,
        ...rows
          .slice(0, REPORTED_ROWS)
          .map(
            (row) =>
              `    ${String(row.files).padStart(5)} files  ${row.invariant} | ${row.message}`,
          ),
        ...(rows.length > REPORTED_ROWS ? [`    …and ${rows.length - REPORTED_ROWS} more`] : []),
      ].join("\n");

const renderSide = ({ label }: Side, report: RespellReport): string =>
  [
    `${label}: ${report.exact} rows re-spelled, ${report.byPrefix} matched through a message the old cap cut`,
    renderRows("observed rows nothing recorded explains", report.appeared),
    renderRows("recorded rows the run no longer produces", report.vanished),
  ]
    .filter((line) => line.length > 0)
    .join("\n");

const baselineRows = (baseline: CorpusBaseline | FamilyBaseline): Row[] =>
  baseline.entries.map(({ invariant, message, files }: CorpusBaselineEntry) => ({
    invariant,
    message,
    files,
  }));

const censusRows = (signatures: readonly CensusSignature[]): Row[] =>
  signatures.map(({ invariant, message, files }) => ({ invariant, message, files }));

type CensusFile = CorpusCensus & { family: FamilyCensus };

const loadCensuses = async (paths: readonly string[]): Promise<CensusFile> => {
  if (paths.length === 0) {
    throw new RespellError({ message: "Pass the census files the nightly produced" });
  }
  const censuses = await Promise.all(
    paths.map(async (file) => (await Bun.file(file).json()) as CensusFile),
  );
  return {
    ...mergeCensuses(censuses),
    family: mergeFamilyCensuses(censuses.map((census) => census.family)),
  };
};

const main = async (paths: string[]): Promise<void> => {
  const census = await loadCensuses(paths);
  const baseline = (await Bun.file(BASELINE_PATH).json()) as CorpusBaseline;
  const sides: { side: Side; recorded: Row[]; observed: Row[] }[] = [
    {
      side: { label: "corpus/baseline.json" },
      recorded: baselineRows(baseline),
      observed: censusRows(census.signatures),
    },
  ];
  for (const family of FAMILY_BASELINE_FAMILIES) {
    // oxlint-disable-next-line no-await-in-loop -- one small file per family, read in a fixed order
    const loaded = await loadFamilyBaseline(family);
    sides.push({
      side: { label: `corpus/baselines/${family}.json` },
      recorded: baselineRows(loaded),
      observed: censusRows(census.family.signatures.filter((row) => row.family === family)),
    });
  }

  const reports = sides.map(({ side, recorded, observed }) => ({
    side,
    report: compareSide(recorded, observed),
  }));
  process.stdout.write(
    `${reports.map(({ side, report }) => renderSide(side, report)).join("\n")}\n`,
  );
  const unexplained = reports.reduce(
    (total, { report }) => total + report.appeared.length + report.vanished.length,
    0,
  );
  if (unexplained > 0) {
    process.stderr.write(
      `${unexplained} row(s) the re-spelling does not explain. Each is a defect class that appeared or stopped being produced, not a change of spelling: decide about it before baking.\n`,
    );
    process.exitCode = 1;
    return;
  }
  process.stdout.write(
    `Every row is accounted for. Bake with \`bun scripts/corpus-gate.ts write-baseline ${paths.join(" ")}\`.\n`,
  );
};

if (import.meta.main) {
  main(process.argv.slice(2)).catch((cause: unknown) => {
    process.stderr.write(`${cause instanceof Error ? cause.message : String(cause)}\n`);
    process.exitCode = 1;
  });
}

export { compareSide, withoutCarriers };
