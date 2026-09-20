/**
 * Whether the committed corpus files could have come from a run.
 *
 * The ratchet answers a different question, and it needs an hour of corpus to
 * answer it: whether the numbers still hold. Nothing asked the cheap question —
 * whether the files are internally consistent at all — so a hand-edited
 * baseline, a row whose `signature` no longer matches its own fields, a
 * duplicate row, a family file holding another family's invariant, a baseline
 * measured over a lock the repository no longer carries, or a report-only
 * exemption for a file the lock dropped could reach `main` with nothing to
 * catch it. Pull-request CI does not download the corpus, and none of these
 * checks needs it.
 *
 * Every check reads committed files alone and is decided by data, never by the
 * clock, a network or a cache.
 */

import path from "node:path";

import type { CorpusBaseline, CorpusBaselineEntry } from "./corpus-baseline";
import { type ExpectedDispositions, validateExpectedDispositions } from "./corpus-dispositions";
import {
  type CorpusInvariantFamily,
  EXTENDED_CORPUS_INVARIANTS,
  familyOf,
  isGatingFamily,
} from "./corpus-invariants/contract";
import {
  FAMILY_BASELINE_FAMILIES,
  type FamilyBaseline,
  familyBaselinePath,
  loadFamilyBaseline,
} from "./corpus-family-baseline";
import {
  BASELINE_PATH,
  CORPUS_DIRECTORY,
  type CorpusLock,
  DEFAULT_CORPUS_TIERS,
  EXPECTED_DISPOSITIONS_PATH,
  EXPECTED_REFUSALS_PATH,
  loadCorpusLock,
} from "./corpus-manifest";
import type { ExpectedRefusals } from "./corpus-refusals";
import {
  deadReportOnlyEntries,
  REPORT_ONLY_FILES_PATH,
  type ReportOnlyFiles,
  reportOnlyFilesDigest,
  validateReportOnlyFiles,
} from "./corpus-report-only";
import { CORPUS_INVARIANTS, failureSignature } from "./corpus-signature";
import { tierScopedLockDigest } from "./corpus-tiers";

const KNOWN_INVARIANTS: ReadonlySet<string> = new Set([
  ...Object.values(CORPUS_INVARIANTS),
  ...Object.values(EXTENDED_CORPUS_INVARIANTS),
]);

/** One thing wrong with a committed file, named where a reader can fix it. */
export type ValidityIssue = { file: string; detail: string };

const shortPath = (absolute: string): string =>
  path.relative(path.dirname(CORPUS_DIRECTORY), absolute);

/** A row as either baseline file writes it; only the family files carry producers. */
type BaselineRow = CorpusBaselineEntry & { producers?: Record<string, number> };

/**
 * A row's `signature` is its three fields joined, and the gate keys everything
 * on it. A row whose signature and fields disagree matches nothing a run
 * produces: it can only ever ratchet down to nothing and be rewritten.
 */
const rowIssues = (
  file: string,
  rows: readonly BaselineRow[],
  family: CorpusInvariantFamily | undefined,
): ValidityIssue[] => {
  const issues: ValidityIssue[] = [];
  const seen = new Set<string>();
  let previous: string | undefined;
  for (const row of rows) {
    if (!KNOWN_INVARIANTS.has(row.invariant)) {
      issues.push({ file, detail: `unknown invariant \`${row.invariant}\`` });
      continue;
    }
    const rebuilt = failureSignature(row);
    if (rebuilt !== row.signature) {
      issues.push({
        file,
        detail: `a signature disagrees with its own fields: ${row.signature} vs ${rebuilt}`,
      });
    }
    if (seen.has(row.signature)) {
      issues.push({ file, detail: `duplicate signature: ${row.signature}` });
    }
    seen.add(row.signature);
    if (previous !== undefined && previous > row.signature) {
      issues.push({ file, detail: `entries are not sorted by signature at ${row.signature}` });
    }
    previous = row.signature;
    if (!Number.isInteger(row.files) || row.files < 1) {
      issues.push({ file, detail: `${row.signature} records ${row.files} files` });
    }
    const owning = familyOf(row.invariant);
    if (!isGatingFamily(owning)) {
      issues.push({
        file,
        detail: `${row.signature} belongs to \`${owning}\`, a report-only family that owns no baseline`,
      });
    }
    if (family !== undefined && owning !== family) {
      issues.push({ file, detail: `${row.signature} belongs to family \`${owning}\`` });
    }
    if (row.producers !== undefined) {
      const counted = Object.values(row.producers).reduce((total, count) => total + count, 0);
      if (counted !== row.files) {
        issues.push({
          file,
          detail: `${row.signature} affects ${row.files} files across ${counted} producer occurrences`,
        });
      }
    }
  }
  return issues;
};

type Digests = { lockDigest: string; reportOnlyDigest: string };

const digestIssues = (file: string, measured: Digests, committed: Digests): ValidityIssue[] => {
  const issues: ValidityIssue[] = [];
  if (measured.lockDigest !== committed.lockDigest) {
    issues.push({
      file,
      detail: `measured over corpus ${measured.lockDigest}; the committed lock is ${committed.lockDigest}. Refetch and rerun with \`write-baseline\`.`,
    });
  }
  if (measured.reportOnlyDigest !== committed.reportOnlyDigest) {
    issues.push({
      file,
      detail: `measured under report-only list ${measured.reportOnlyDigest}; the committed list is ${committed.reportOnlyDigest}.`,
    });
  }
  return issues;
};

export type CommittedCorpusFiles = {
  baseline: CorpusBaseline;
  families: ReadonlyMap<CorpusInvariantFamily, FamilyBaseline>;
  refusals: ExpectedRefusals;
  dispositions: ExpectedDispositions;
  reportOnly: ReportOnlyFiles;
  lock: CorpusLock;
};

/**
 * Every internal inconsistency in the committed corpus files.
 *
 * Pure over its input, so the same files always produce the same list and a
 * test can hand it a broken one without writing anything to disk.
 */
export const corpusValidityIssues = ({
  baseline,
  families,
  refusals,
  dispositions,
  reportOnly,
  lock,
}: CommittedCorpusFiles): ValidityIssue[] => {
  const issues: ValidityIssue[] = [];
  const baselineFile = shortPath(BASELINE_PATH);
  // The committed baselines are the ones pull-request CI and the nightly
  // ratchet, and both run the default tier selection.
  const committed: Digests = {
    lockDigest: tierScopedLockDigest(lock, DEFAULT_CORPUS_TIERS),
    reportOnlyDigest: reportOnlyFilesDigest(reportOnly),
  };

  issues.push(...digestIssues(baselineFile, baseline, committed));
  issues.push(...rowIssues(baselineFile, baseline.entries, undefined));
  if (!Number.isInteger(baseline.failedFiles) || baseline.failedFiles < 0) {
    issues.push({ file: baselineFile, detail: `failedFiles is ${baseline.failedFiles}` });
  }

  for (const family of FAMILY_BASELINE_FAMILIES) {
    const file = shortPath(familyBaselinePath(family));
    const loaded = families.get(family);
    if (loaded === undefined) {
      issues.push({ file, detail: "a gating family has no baseline file" });
      continue;
    }
    if (loaded.family !== family) {
      issues.push({ file, detail: `declares family \`${loaded.family}\`` });
    }
    issues.push(...digestIssues(file, loaded, committed));
    issues.push(...rowIssues(file, loaded.entries, family));
    if (loaded.failedFiles > loaded.files) {
      issues.push({ file, detail: `${loaded.failedFiles} of ${loaded.files} files fail` });
    }
  }

  // `corpus/baseline.json` aggregates every invariant, so a family row lives in
  // two committed files and the two have to agree about it.
  const aggregated = new Map(baseline.entries.map((entry) => [entry.signature, entry.files]));
  for (const [family, loaded] of families) {
    for (const entry of loaded.entries) {
      const inCore = aggregated.get(entry.signature);
      if (inCore !== undefined && inCore !== entry.files) {
        issues.push({
          file: shortPath(familyBaselinePath(family)),
          detail: `${entry.signature} affects ${entry.files} files here and ${inCore} in ${shortPath(BASELINE_PATH)}`,
        });
      }
    }
  }

  const refusalsFile = shortPath(EXPECTED_REFUSALS_PATH);
  const seenRefusals = new Set<string>();
  for (const entry of refusals.entries) {
    if (seenRefusals.has(entry.signature)) {
      issues.push({ file: refusalsFile, detail: `duplicate signature: ${entry.signature}` });
    }
    seenRefusals.add(entry.signature);
    if (entry.reason.trim().length === 0) {
      issues.push({ file: refusalsFile, detail: `${entry.signature} states no reason` });
    }
    // A refusal is a decision and a baseline row is a defect. A signature in
    // both would be excused and ratcheted at once.
    if (aggregated.has(entry.signature)) {
      issues.push({ file: refusalsFile, detail: `${entry.signature} is also a baseline defect` });
    }
  }

  const dispositionsFile = shortPath(EXPECTED_DISPOSITIONS_PATH);
  for (const detail of validateExpectedDispositions(dispositions, refusals)) {
    issues.push({ file: dispositionsFile, detail });
  }
  // A pattern restricted to an invariant nothing produces claims nothing, and
  // would read as a live decision until the next full run says otherwise.
  for (const entry of dispositions.entries) {
    if (entry.match.kind !== "path") {
      continue;
    }
    for (const invariant of entry.match.invariants) {
      if (!KNOWN_INVARIANTS.has(invariant)) {
        issues.push({
          file: dispositionsFile,
          detail: `${entry.id}: unknown invariant \`${invariant}\``,
        });
      }
    }
  }

  const reportOnlyFile = shortPath(REPORT_ONLY_FILES_PATH);
  for (const detail of validateReportOnlyFiles(reportOnly)) {
    issues.push({ file: reportOnlyFile, detail });
  }
  // An exemption is granted for a file, not for a path: a repin that replaced
  // the bytes has to be reviewed again rather than inheriting the decision.
  for (const detail of deadReportOnlyEntries(reportOnly, lock)) {
    issues.push({ file: reportOnlyFile, detail });
  }

  return issues;
};

/** Read every committed file the checks read. */
export const loadCommittedCorpusFiles = async (): Promise<CommittedCorpusFiles> => {
  const [baseline, refusals, dispositions, reportOnly, lock] = await Promise.all([
    Bun.file(BASELINE_PATH).json() as Promise<CorpusBaseline>,
    Bun.file(EXPECTED_REFUSALS_PATH).json() as Promise<ExpectedRefusals>,
    Bun.file(EXPECTED_DISPOSITIONS_PATH).json() as Promise<ExpectedDispositions>,
    Bun.file(REPORT_ONLY_FILES_PATH).json() as Promise<ReportOnlyFiles>,
    loadCorpusLock(),
  ]);
  const loaded = await Promise.all(
    FAMILY_BASELINE_FAMILIES.map(
      async (family) => [family, await loadFamilyBaseline(family)] as const,
    ),
  );
  return { baseline, families: new Map(loaded), refusals, dispositions, reportOnly, lock };
};

export const renderValidityIssues = (issues: readonly ValidityIssue[]): string =>
  issues.map(({ file, detail }) => `- ${file}: ${detail}`).join("\n");
