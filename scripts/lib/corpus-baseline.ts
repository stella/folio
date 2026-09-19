/**
 * The shrink-only baseline.
 *
 * The corpus finds more defects than anyone will fix in one change, so the gate
 * records what fails today and refuses anything worse. Like the repository's
 * other ratchets, the recorded numbers may only move one way: a signature that
 * gains files is a regression, a signature the corpus has never seen is a new
 * defect, and a signature that improved must be rewritten down before the gate
 * passes again, so an accidental re-regression cannot hide inside old slack.
 *
 * The baseline is bound to the lock digest of the corpus it was measured over.
 * Repinning a source changes which files exist, which changes the counts, so a
 * corpus change forces an explicit refresh instead of silently redefining what
 * "no worse" means.
 */

import type { CorpusCensus } from "./corpus-census";
import { isGatingFailure } from "./corpus-invariants/contract";
import type { CorpusInvariant } from "./corpus-signature";

/**
 * `corpus/baseline.json` aggregates every invariant, so the report-only
 * families reach it too and would ratchet here even after giving up their own
 * baseline file. They are dropped on the way in and on the way out, so the two
 * sides always compare the same set.
 */
const gatingSignatures = <T extends { invariant: CorpusInvariant }>(
  signatures: readonly T[],
): T[] => signatures.filter(isGatingFailure);

export type CorpusBaselineEntry = {
  signature: string;
  invariant: CorpusInvariant;
  message: string;
  frame: string;
  files: number;
};

export type CorpusBaseline = {
  schemaVersion: 1;
  lockDigest: string;
  /** The report-only list these counts were measured under. */
  reportOnlyDigest: string;
  failedFiles: number;
  entries: CorpusBaselineEntry[];
};

/**
 * Whether a run measured the set of files it was supposed to.
 *
 * A truncated file contributes no gating evidence, so every truncation moves
 * the boundary of what was compared. The files that are allowed to sit outside
 * that boundary are named in `corpus/report-only-files.json` and never reach
 * this count, so any truncation left here is a file whose evidence went
 * missing for a reason nobody wrote down: the run is degraded, whatever share
 * of the corpus it is.
 */
export const isDegradedRun = (census: CorpusCensus): boolean => (census.truncated ?? 0) > 0;

/** How a degraded run names what it lost, and what to do about it. */
export const describeDegradedRun = (census: CorpusCensus): string => {
  const named = (census.truncatedExamples ?? [])
    .map(({ file, stage }) => `${file.sourceId}/${file.path} (stopped at ${stage})`)
    .join(", ");
  return `${census.truncated} of ${census.files} files stopped at a budget without being listed as report-only: ${named}. Add each to corpus/report-only-files.json with a reason, or make it finish inside the budget.`;
};

export const baselineFromCensus = (census: CorpusCensus): CorpusBaseline => ({
  schemaVersion: 1,
  lockDigest: census.lockDigest,
  reportOnlyDigest: census.reportOnlyDigest,
  failedFiles: census.failedFiles,
  entries: gatingSignatures(census.signatures)
    .map(({ signature, invariant, message, frame, files }) => ({
      signature,
      invariant,
      message,
      frame,
      files,
    }))
    .sort((left, right) => (left.signature < right.signature ? -1 : 1)),
});

export type BaselineViolation = {
  kind:
    | "new-signature"
    | "more-files"
    | "fewer-files"
    | "resolved-signature"
    | "corpus-changed"
    | "run-degraded"
    | "report-only-list-changed";
  signature: string;
  detail: string;
};

/**
 * Findings that report what a run was not asked to see, rather than what it saw.
 *
 * Listing a file as report-only removes its findings from the comparison, so
 * against a baseline written before the listing every signature that file
 * carried looks smaller or gone. That is the listing working, not a shrink to
 * ratchet, and reporting it as resolved would erase a defect the corpus still
 * has. Growth is unaffected and still fails: a listing can only remove
 * evidence. These are printed and clear on the next `write-baseline`.
 */
const INFORMATIONAL_KINDS: ReadonlySet<BaselineViolation["kind"]> = new Set([
  "report-only-list-changed",
]);

export const isFailingViolation = (violation: { kind: string }): boolean =>
  !INFORMATIONAL_KINDS.has(violation.kind as BaselineViolation["kind"]);

/**
 * Whether the baseline was measured under the report-only list this run read.
 *
 * Every baseline file records the digest, the family ones included, for the
 * reason each already records the lock digest: a file is only comparable
 * against a run that measured the same corpus in the same way.
 */
export const reportOnlyListChanged = (
  baseline: { reportOnlyDigest: string },
  census: { reportOnlyDigest: string },
): boolean => baseline.reportOnlyDigest !== census.reportOnlyDigest;

export const compareToBaseline = (
  baseline: CorpusBaseline,
  census: CorpusCensus,
): BaselineViolation[] => {
  if (baseline.lockDigest !== census.lockDigest) {
    return [
      {
        kind: "corpus-changed",
        signature: "-",
        detail: `the baseline was measured over corpus ${baseline.lockDigest.slice(0, 12)}, this run saw ${census.lockDigest.slice(0, 12)}; rerun with \`write-baseline\``,
      },
    ];
  }

  // A file that stopped at a budget without being listed took its evidence
  // with it, so the comparison is over a set nobody chose: there is nothing
  // sound to report until the run is repeated or the file is listed.
  if (isDegradedRun(census)) {
    return [{ kind: "run-degraded", signature: "-", detail: describeDegradedRun(census) }];
  }

  const listChanged = reportOnlyListChanged(baseline, census);
  const recorded = new Map(
    gatingSignatures(baseline.entries).map((entry) => [entry.signature, entry]),
  );
  const violations: BaselineViolation[] = [];
  for (const observed of gatingSignatures(census.signatures)) {
    const entry = recorded.get(observed.signature);
    if (entry === undefined) {
      const example = observed.examples.at(0);
      violations.push({
        kind: "new-signature",
        signature: observed.signature,
        detail: `${observed.files} file(s), e.g. ${example === undefined ? "-" : `${example.sourceId}/${example.path}`}`,
      });
      continue;
    }
    recorded.delete(observed.signature);
    if (observed.files > entry.files) {
      violations.push({
        kind: "more-files",
        signature: observed.signature,
        detail: `${observed.files} files fail, the baseline allows ${entry.files}`,
      });
      continue;
    }
    if (observed.files < entry.files) {
      violations.push(
        listChanged
          ? {
              kind: "report-only-list-changed",
              signature: observed.signature,
              detail: `${observed.files} files fail, down from ${entry.files}, but the report-only list changed since this baseline was written; kept`,
            }
          : {
              kind: "fewer-files",
              signature: observed.signature,
              detail: `${observed.files} files fail, down from ${entry.files}; rerun with \`write-baseline\``,
            },
      );
    }
  }
  for (const entry of recorded.values()) {
    violations.push(
      listChanged
        ? {
            kind: "report-only-list-changed",
            signature: entry.signature,
            detail:
              "not seen this run, but the report-only list changed since this baseline was written; kept",
          }
        : {
            kind: "resolved-signature",
            signature: entry.signature,
            detail: "no longer fails; remove it with `write-baseline`",
          },
    );
  }
  return violations;
};

/** Every ratchet this gate runs reports the same shape, so one renderer serves. */
export type RatchetViolation = { kind: string; signature: string; detail: string };

export const renderViolations = (violations: readonly RatchetViolation[]): string =>
  violations
    .map(({ kind, signature, detail }) => `- [${kind}] ${signature}\n    ${detail}`)
    .join("\n");
