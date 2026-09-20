/**
 * The census of the extended invariant families.
 *
 * It sits beside `corpus-census.ts` rather than inside it: the original census
 * is the contract `corpus/baseline.json` and the CI workflow are written
 * against, and a family census carries facts that one does not, above all which
 * producers a signature fires on. Two censuses merge independently, so a
 * sharded run still reduces to one report.
 *
 * Whether a file is a performance outlier is a fact about the whole corpus, so
 * the builder records per-file cost and the verdict is taken once, at merge
 * time, from the median. Per-stage timings are kept only as a bounded top list:
 * the slowest twenty files per stage is the whole of what a reader does with
 * them, and keeping every timing would put a megabyte of noise in the report.
 */

import {
  type CorpusInvariantFamily,
  EXTENDED_INVARIANT_FAMILY,
  familyOf,
  isGatingFailure,
} from "./corpus-invariants/contract";
import {
  CORPUS_EVIDENCE,
  type CorpusEvidence,
  type CorpusFileId,
  MAX_EXAMPLES_PER_SIGNATURE,
} from "./corpus-census";
import {
  type CorpusFailure,
  type CorpusInvariant,
  distinctBySignature,
  failureSignature,
} from "./corpus-signature";

export const SLOWEST_FILES_PER_STAGE = 20;

/** Producer label to the number of files it accounts for. */
export type ProducerCounts = Record<string, number>;

export type FamilySignature = {
  signature: string;
  family: CorpusInvariantFamily;
  invariant: CorpusInvariant;
  message: string;
  frame: string;
  files: number;
  producers: ProducerCounts;
  examples: CorpusFileId[];
};

export type FamilyTotals = { files: number; failedFiles: number };

export type StageTiming = { file: CorpusFileId; bytes: number; ms: number };

/** One file's parse cost, from which the outlier verdict is taken at merge time. */
export type FileCost = {
  file: CorpusFileId;
  bytes: number;
  parseMs: number;
  peakRssBytes: number;
  /** What the fixed reference package cost beside this file, so load divides out. */
  referenceMs: number;
  /** Carried so the outlier verdict, taken later, can still name its producers. */
  producer: string;
};

export type FamilyCensus = {
  schemaVersion: 1;
  lockDigest: string;
  /** The `corpus/report-only-files.json` this run read, as the core census records it. */
  reportOnlyDigest: string;
  files: number;
  producers: ProducerCounts;
  totals: Record<string, FamilyTotals>;
  signatures: FamilySignature[];
  /** Stage name to its slowest files, longest first. */
  slowest: Record<string, StageTiming[]>;
  costs: FileCost[];
};

const EXTENDED_FAMILIES = Object.values(EXTENDED_INVARIANT_FAMILY);

export { familyOf } from "./corpus-invariants/contract";

export const emptyFamilyCensus = (lockDigest: string, reportOnlyDigest: string): FamilyCensus => ({
  schemaVersion: 1,
  lockDigest,
  reportOnlyDigest,
  files: 0,
  producers: {},
  totals: {},
  signatures: [],
  slowest: {},
  costs: [],
});

const bump = (counts: ProducerCounts, key: string): void => {
  counts[key] = (counts[key] ?? 0) + 1;
};

/**
 * The failures one file contributes, at most one per signature.
 *
 * A file that carries no gating evidence keeps its timing findings, which are
 * the whole story of why it stopped, and contributes nothing to the families
 * that gate. Exported because the per-file signature output has to name the
 * same set this census counted: two spellings of the rule would let an
 * attribution disagree with the count it is supposed to explain.
 */
export const countedFailures = (
  failures: readonly CorpusFailure[],
  evidence: CorpusEvidence,
): CorpusFailure[] =>
  distinctBySignature(
    evidence === CORPUS_EVIDENCE.gating
      ? failures
      : failures.filter((failure) => !isGatingFailure(failure)),
  );

const keepSlowest = (into: StageTiming[], timing: StageTiming): void => {
  into.push(timing);
  into.sort((left, right) => right.ms - left.ms);
  into.length = Math.min(into.length, SLOWEST_FILES_PER_STAGE);
};

export type ObservedFile = {
  file: CorpusFileId;
  bytes: number;
  parseMs: number;
  peakRssBytes: number;
  referenceMs: number;
  producer: string;
  failures: readonly CorpusFailure[];
  timings: Readonly<Record<string, number>>;
  /**
   * Whether this file's findings may gate, as `evidenceOf` decided it.
   *
   * Stated rather than defaulted: the core census and this one must agree
   * about every file, and a default here would let them drift apart.
   */
  evidence: CorpusEvidence;
};

export class FamilyCensusBuilder {
  readonly #census: FamilyCensus;
  readonly #bySignature = new Map<string, FamilySignature>();

  constructor(lockDigest: string, reportOnlyDigest: string) {
    this.#census = emptyFamilyCensus(lockDigest, reportOnlyDigest);
  }

  add({
    file,
    bytes,
    parseMs,
    peakRssBytes,
    referenceMs,
    producer,
    failures,
    timings,
    evidence,
  }: ObservedFile): void {
    this.#census.files += 1;
    bump(this.#census.producers, producer);
    this.#census.costs.push({ file, bytes, parseMs, peakRssBytes, referenceMs, producer });

    for (const [stage, ms] of Object.entries(timings)) {
      const into = (this.#census.slowest[stage] ??= []);
      keepSlowest(into, { file, bytes, ms });
    }

    const counted = countedFailures(failures, evidence);

    const familiesTouched = new Set<CorpusInvariantFamily>();
    for (const failure of counted) {
      const family = familyOf(failure.invariant);
      familiesTouched.add(family);
      const signature = failureSignature(failure);
      const existing = this.#bySignature.get(signature);
      if (existing === undefined) {
        this.#bySignature.set(signature, {
          signature,
          family,
          invariant: failure.invariant,
          message: failure.message,
          frame: failure.frame,
          files: 1,
          producers: { [producer]: 1 },
          examples: [file],
        });
        continue;
      }
      existing.files += 1;
      bump(existing.producers, producer);
      if (existing.examples.length < MAX_EXAMPLES_PER_SIGNATURE) {
        existing.examples.push(file);
      }
    }

    for (const family of EXTENDED_FAMILIES) {
      const totals = (this.#census.totals[family] ??= { files: 0, failedFiles: 0 });
      totals.files += 1;
      if (familiesTouched.has(family)) {
        totals.failedFiles += 1;
      }
    }
  }

  build(): FamilyCensus {
    return { ...this.#census, signatures: sortSignatures([...this.#bySignature.values()]) };
  }
}

/** Busiest first, then by signature so a census is byte-stable across runs. */
const compareByFilesThenSignature = (left: FamilySignature, right: FamilySignature): number => {
  if (left.files !== right.files) {
    return right.files - left.files;
  }
  return left.signature < right.signature ? -1 : 1;
};

const sortSignatures = (signatures: FamilySignature[]): FamilySignature[] =>
  signatures.sort(compareByFilesThenSignature);

const mergeCounts = (into: ProducerCounts, from: ProducerCounts): void => {
  for (const [key, count] of Object.entries(from)) {
    into[key] = (into[key] ?? 0) + count;
  }
};

export const mergeFamilyCensuses = (censuses: readonly FamilyCensus[]): FamilyCensus => {
  const first = censuses.at(0);
  if (first === undefined) {
    throw new Error("A family census merge needs at least one census");
  }
  const merged = emptyFamilyCensus(first.lockDigest, first.reportOnlyDigest);
  const bySignature = new Map<string, FamilySignature>();
  for (const census of censuses) {
    merged.files += census.files;
    mergeCounts(merged.producers, census.producers);
    merged.costs.push(...census.costs);
    for (const [family, totals] of Object.entries(census.totals)) {
      const into = (merged.totals[family] ??= { files: 0, failedFiles: 0 });
      into.files += totals.files;
      into.failedFiles += totals.failedFiles;
    }
    for (const [stage, timings] of Object.entries(census.slowest)) {
      const into = (merged.slowest[stage] ??= []);
      for (const timing of timings) {
        keepSlowest(into, timing);
      }
    }
    for (const signature of census.signatures) {
      const existing = bySignature.get(signature.signature);
      if (existing === undefined) {
        bySignature.set(signature.signature, {
          ...signature,
          producers: { ...signature.producers },
          examples: [...signature.examples],
        });
        continue;
      }
      existing.files += signature.files;
      mergeCounts(existing.producers, signature.producers);
      existing.examples.push(
        ...signature.examples.slice(0, MAX_EXAMPLES_PER_SIGNATURE - existing.examples.length),
      );
    }
  }
  merged.signatures = sortSignatures([...bySignature.values()]);
  return merged;
};

export type LateFailure = {
  file: CorpusFileId;
  producer: string;
  failures: readonly CorpusFailure[];
};

/**
 * Fold in failures that could only be decided once the whole run was in.
 *
 * Whether a file is a performance outlier is a fact about the corpus, not about
 * the file, so the verdict cannot be taken in the worker. Merging it here keeps
 * the performance family reporting through the same signatures, producers and
 * baseline machinery as every other family.
 */
export const censusWithLateFailures = (
  census: FamilyCensus,
  late: readonly LateFailure[],
): FamilyCensus => {
  const bySignature = new Map(
    census.signatures.map((signature) => [
      signature.signature,
      { ...signature, producers: { ...signature.producers }, examples: [...signature.examples] },
    ]),
  );
  const totals: Record<string, FamilyTotals> = Object.fromEntries(
    Object.entries(census.totals).map(([family, value]) => [family, { ...value }]),
  );
  const failedByFamily = new Map<CorpusInvariantFamily, Set<string>>();

  for (const { file, producer, failures } of late) {
    for (const failure of distinctBySignature(failures)) {
      const family = familyOf(failure.invariant);
      const signature = failureSignature(failure);
      const existing = bySignature.get(signature);
      if (existing === undefined) {
        bySignature.set(signature, {
          signature,
          family,
          invariant: failure.invariant,
          message: failure.message,
          frame: failure.frame,
          files: 1,
          producers: { [producer]: 1 },
          examples: [file],
        });
      } else {
        existing.files += 1;
        bump(existing.producers, producer);
        if (existing.examples.length < MAX_EXAMPLES_PER_SIGNATURE) {
          existing.examples.push(file);
        }
      }
      const failed = failedByFamily.get(family) ?? new Set<string>();
      failed.add(file.sha256);
      failedByFamily.set(family, failed);
    }
  }

  for (const [family, failed] of failedByFamily) {
    const into = (totals[family] ??= { files: census.files, failedFiles: 0 });
    into.failedFiles += failed.size;
  }
  return { ...census, totals, signatures: sortSignatures([...bySignature.values()]) };
};

const PRODUCERS_PER_SIGNATURE = 4;

export type RenderFamilyCensusOptions = {
  census: FamilyCensus;
  signaturesPerFamily: number;
  slowestPerStage: number;
};

export const renderFamilyCensus = ({
  census,
  signaturesPerFamily,
  slowestPerStage,
}: RenderFamilyCensusOptions): string => {
  const lines = [`producers: ${describeProducers(census.producers, 12)}`];
  for (const family of [...new Set(EXTENDED_FAMILIES)].sort()) {
    const totals = census.totals[family] ?? { files: 0, failedFiles: 0 };
    const signatures = census.signatures.filter((signature) => signature.family === family);
    lines.push(
      `${family}: ${totals.files} files, ${totals.files - totals.failedFiles} pass, ${totals.failedFiles} fail, ${signatures.length} signatures`,
    );
    for (const signature of signatures.slice(0, signaturesPerFamily)) {
      const example = signature.examples.at(0);
      lines.push(
        `  ${String(signature.files).padStart(5)}  ${signature.message} | ${signature.frame}`,
        `         producers: ${describeProducers(signature.producers, PRODUCERS_PER_SIGNATURE)}`,
        `         e.g. ${example === undefined ? "-" : `${example.sourceId}/${example.path}`}`,
      );
    }
  }
  for (const stage of Object.keys(census.slowest).sort()) {
    const timings = (census.slowest[stage] ?? []).slice(0, slowestPerStage);
    if (timings.length === 0) {
      continue;
    }
    lines.push(`slowest ${stage}:`);
    for (const { file, bytes, ms } of timings) {
      lines.push(
        `  ${ms.toFixed(0).padStart(7)}ms  ${(bytes / 1_048_576).toFixed(2).padStart(7)}MB  ${file.sourceId}/${file.path}`,
      );
    }
  }
  return lines.join("\n");
};

/** The producers a signature fires on, busiest first, as a census line reads them. */
const compareProducerEntries = (
  [leftLabel, left]: readonly [string, number],
  [rightLabel, right]: readonly [string, number],
): number => {
  if (left !== right) {
    return right - left;
  }
  return leftLabel < rightLabel ? -1 : 1;
};

export const describeProducers = (producers: ProducerCounts, limit: number): string =>
  Object.entries(producers)
    .sort(compareProducerEntries)
    .slice(0, limit)
    .map(([label, count]) => `${label} ${count}`)
    .join(", ");
