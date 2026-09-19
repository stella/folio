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
import type { CorpusInvariant } from "./corpus-signature";

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
  entries: census.signatures
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
  kind: "new-signature" | "more-files" | "fewer-files" | "resolved-signature" | "corpus-changed";
  signature: string;
  detail: string;
};

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

  const recorded = new Map(baseline.entries.map((entry) => [entry.signature, entry]));
  const violations: BaselineViolation[] = [];
  for (const observed of census.signatures) {
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
      violations.push({
        kind: "fewer-files",
        signature: observed.signature,
        detail: `${observed.files} files fail, down from ${entry.files}; rerun with \`write-baseline\``,
      });
    }
  }
  for (const entry of recorded.values()) {
    violations.push({
      kind: "resolved-signature",
      signature: entry.signature,
      detail: `no longer fails; remove it with \`write-baseline\``,
    });
  }
  return violations;
};

/** Every ratchet this gate runs reports the same shape, so one renderer serves. */
export type RatchetViolation = { kind: string; signature: string; detail: string };

export const renderViolations = (violations: readonly RatchetViolation[]): string =>
  violations
    .map(({ kind, signature, detail }) => `- [${kind}] ${signature}\n    ${detail}`)
    .join("\n");
