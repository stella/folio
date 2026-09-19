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

import { type CorpusCensus, MAX_TRUNCATED_FRACTION } from "./corpus-census";
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
  failedFiles: number;
  entries: CorpusBaselineEntry[];
};

export const baselineFromCensus = (census: CorpusCensus): CorpusBaseline => ({
  schemaVersion: 1,
  lockDigest: census.lockDigest,
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
    | "unobserved-truncated";
  signature: string;
  detail: string;
};

/**
 * Findings that report what a run could not see, rather than what it saw.
 *
 * Truncation only ever removes evidence: a file that stopped at a budget can
 * make a signature look smaller or gone, never bigger or new. So a shrink
 * measured by a run that truncated anything is a shrink that may not be real,
 * and demanding the baseline be written down to it would bake the truncation
 * in. These are printed and do not fail the gate.
 */
const INFORMATIONAL_KINDS: ReadonlySet<BaselineViolation["kind"]> = new Set([
  "unobserved-truncated",
]);

export const isFailingViolation = (violation: { kind: string }): boolean =>
  !INFORMATIONAL_KINDS.has(violation.kind as BaselineViolation["kind"]);

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

  // Past this share of the corpus the run has not measured enough to be
  // compared at all: too many files stopped early for "no new signature" to
  // mean anything.
  const truncated = census.truncated ?? 0;
  if (truncated > census.files * MAX_TRUNCATED_FRACTION) {
    return [
      {
        kind: "run-degraded",
        signature: "-",
        detail: `${truncated} of ${census.files} files stopped at a budget (over ${MAX_TRUNCATED_FRACTION * 100}%); the run is too thin to compare, rerun it`,
      },
    ];
  }

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
        truncated
          ? {
              kind: "unobserved-truncated",
              signature: observed.signature,
              detail: `${observed.files} files fail, down from ${entry.files}, but ${truncated} file(s) stopped at a budget this run`,
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
      truncated
        ? {
            kind: "unobserved-truncated",
            signature: entry.signature,
            detail: `not seen this run, but ${truncated} file(s) stopped at a budget; kept`,
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
