/**
 * Refusals the corpus is expected to produce.
 *
 * Not every failing file is a defect. Some are packages folio deliberately
 * declines: a part nested past the depth bound, markup the resource preflight
 * cannot scan safely. Those are typed errors with a stated reason, and leaving
 * them in `corpus/baseline.json` makes the defect list read as longer than it
 * is and hides a real regression among entries nobody intends to fix.
 *
 * So the census partitions on this list. A signature it names is reported as a
 * refusal with its reason; everything else ratchets against the baseline. The
 * list is hand-written: a `reason` is a decision a person makes, and promoting
 * a signature is never something a rewrite does on its own. `write-baseline`
 * only refreshes the counts, which ratchet the same way the baseline does, so
 * a refusal that starts affecting more files is still a finding and one that
 * stops reproducing has to be removed.
 */

import type { CorpusCensus, CensusSignature } from "./corpus-census";

export type ExpectedRefusalEntry = {
  signature: string;
  /** Why folio declines this input, in a sentence a reviewer can weigh. */
  reason: string;
  files: number;
};

export type ExpectedRefusals = {
  schemaVersion: 1;
  entries: ExpectedRefusalEntry[];
};

export type PartitionedCensus = {
  /** The census the baseline ratchets against: signatures nobody allowlisted. */
  defects: CorpusCensus;
  refusals: CensusSignature[];
};

export const partitionExpectedRefusals = (
  census: CorpusCensus,
  refusals: ExpectedRefusals,
): PartitionedCensus => {
  const expected = new Set(refusals.entries.map((entry) => entry.signature));
  const defectSignatures: CensusSignature[] = [];
  const refusalSignatures: CensusSignature[] = [];
  for (const signature of census.signatures) {
    (expected.has(signature.signature) ? refusalSignatures : defectSignatures).push(signature);
  }
  return {
    defects: { ...census, signatures: defectSignatures },
    refusals: refusalSignatures,
  };
};

/**
 * The list with every count refreshed from what the run observed.
 *
 * Reasons and the set of allowlisted signatures are preserved exactly: this
 * refreshes a ratchet, it does not decide what belongs on the list.
 */
export const refreshedExpectedRefusals = (
  refusals: ExpectedRefusals,
  observed: readonly CensusSignature[],
): ExpectedRefusals => {
  const filesBySignature = new Map(observed.map((entry) => [entry.signature, entry.files]));
  return {
    schemaVersion: 1,
    entries: refusals.entries
      .map((entry) => ({ ...entry, files: filesBySignature.get(entry.signature) ?? 0 }))
      .sort((left, right) => (left.signature < right.signature ? -1 : 1)),
  };
};

export type ExpectedRefusalViolation = {
  kind: "more-files" | "fewer-files" | "resolved-refusal";
  signature: string;
  detail: string;
};

export const compareToExpectedRefusals = (
  refusals: ExpectedRefusals,
  observed: readonly CensusSignature[],
): ExpectedRefusalViolation[] => {
  const filesBySignature = new Map(observed.map((entry) => [entry.signature, entry.files]));
  const violations: ExpectedRefusalViolation[] = [];
  for (const entry of refusals.entries) {
    const files = filesBySignature.get(entry.signature) ?? 0;
    if (files > entry.files) {
      violations.push({
        kind: "more-files",
        signature: entry.signature,
        detail: `${files} files are refused, the list allows ${entry.files}`,
      });
      continue;
    }
    if (files === 0) {
      violations.push({
        kind: "resolved-refusal",
        signature: entry.signature,
        detail:
          "nothing is refused this way any more; remove it from corpus/expected-refusals.json",
      });
      continue;
    }
    if (files < entry.files) {
      violations.push({
        kind: "fewer-files",
        signature: entry.signature,
        detail: `${files} files are refused, down from ${entry.files}; rerun with \`write-baseline\``,
      });
    }
  }
  return violations;
};

export const renderExpectedRefusals = (
  refusals: ExpectedRefusals,
  observed: readonly CensusSignature[],
): string => {
  if (refusals.entries.length === 0) {
    return "";
  }
  const filesBySignature = new Map(observed.map((entry) => [entry.signature, entry.files]));
  return [
    `  expected refusals ${refusals.entries.length}:`,
    ...refusals.entries.flatMap((entry) => [
      `  ${String(filesBySignature.get(entry.signature) ?? 0).padStart(5)}  ${entry.signature}`,
      `         ${entry.reason}`,
    ]),
  ].join("\n");
};
