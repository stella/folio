/**
 * One shrink-only baseline file per invariant family.
 *
 * The ratchet rule is the one `corpus-baseline.ts` states, and the reasons are
 * the same. What changes here is ownership: a family that gains or loses
 * findings rewrites its own file and no other, so re-measuring the schema
 * validator cannot conflict with a branch re-measuring parse tolerance, and an
 * invariant added later arrives as a new file rather than a rewrite of a shared
 * one.
 *
 * Each file is bound to the tier-scoped digest of the corpus it was measured
 * over, so running a different tier selection is a refusal rather than a
 * silent comparison against the wrong corpus.
 */

import path from "node:path";

import { TaggedError } from "better-result";

import type { BaselineViolation } from "./corpus-baseline";
import { type FamilyCensus, type FamilySignature } from "./corpus-family-census";
import {
  type CorpusInvariantFamily,
  EXTENDED_INVARIANT_FAMILY,
  isGatingFamily,
} from "./corpus-invariants/contract";
import { CORPUS_DIRECTORY, writeJsonFile } from "./corpus-manifest";
import type { CorpusInvariant } from "./corpus-signature";

export class CorpusFamilyBaselineError extends TaggedError("CorpusFamilyBaselineError")<{
  message: string;
}> {}

export const FAMILY_BASELINE_DIRECTORY = path.join(CORPUS_DIRECTORY, "baselines");

export const familyBaselinePath = (family: CorpusInvariantFamily): string =>
  path.join(FAMILY_BASELINE_DIRECTORY, `${family}.json`);

/**
 * The families that own a file here: every extended one that gates, never
 * `core` and never a report-only family, which is measured but not compared.
 */
export const FAMILY_BASELINE_FAMILIES: readonly CorpusInvariantFamily[] = [
  ...new Set(Object.values(EXTENDED_INVARIANT_FAMILY)),
]
  .filter((family) => isGatingFamily(family))
  .sort();

export type FamilyBaselineEntry = {
  signature: string;
  invariant: CorpusInvariant;
  message: string;
  frame: string;
  files: number;
  /** Which producers trigger it, so a reader knows whether it is a quirk or a format defect. */
  producers: Record<string, number>;
};

export type FamilyBaseline = {
  schemaVersion: 1;
  family: CorpusInvariantFamily;
  lockDigest: string;
  files: number;
  failedFiles: number;
  entries: FamilyBaselineEntry[];
};

const signaturesOf = (census: FamilyCensus, family: CorpusInvariantFamily): FamilySignature[] =>
  census.signatures.filter((signature) => signature.family === family);

export const familyBaselineFromCensus = (
  census: FamilyCensus,
  family: CorpusInvariantFamily,
): FamilyBaseline => ({
  schemaVersion: 1,
  family,
  lockDigest: census.lockDigest,
  files: census.totals[family]?.files ?? 0,
  failedFiles: census.totals[family]?.failedFiles ?? 0,
  entries: signaturesOf(census, family)
    .map(({ signature, invariant, message, frame, files, producers }) => ({
      signature,
      invariant,
      message,
      frame,
      files,
      producers,
    }))
    .sort((left, right) => (left.signature < right.signature ? -1 : 1)),
});

export const compareFamilyToBaseline = (
  baseline: FamilyBaseline,
  census: FamilyCensus,
): BaselineViolation[] => {
  if (baseline.lockDigest !== census.lockDigest) {
    return [
      {
        kind: "corpus-changed",
        signature: "-",
        detail: `${baseline.family} was measured over corpus ${baseline.lockDigest.slice(0, 16)}, this run saw ${census.lockDigest.slice(0, 16)}; rerun with \`write-baseline\``,
      },
    ];
  }

  const recorded = new Map(baseline.entries.map((entry) => [entry.signature, entry]));
  const violations: BaselineViolation[] = [];
  for (const observed of signaturesOf(census, baseline.family)) {
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
      detail: "no longer fails; remove it with `write-baseline`",
    });
  }
  return violations;
};

export const loadFamilyBaseline = async (
  family: CorpusInvariantFamily,
): Promise<FamilyBaseline> => {
  const file = Bun.file(familyBaselinePath(family));
  if (!(await file.exists())) {
    throw new CorpusFamilyBaselineError({
      message: `corpus/baselines/${family}.json is missing. Generate it with \`corpus-gate.ts write-baseline\`.`,
    });
  }
  return (await file.json()) as FamilyBaseline;
};

export const writeFamilyBaselines = async (census: FamilyCensus): Promise<void> => {
  await Promise.all(
    FAMILY_BASELINE_FAMILIES.map((family) =>
      writeJsonFile(familyBaselinePath(family), familyBaselineFromCensus(census, family)),
    ),
  );
};
