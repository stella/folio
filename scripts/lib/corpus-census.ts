/**
 * The census: what a corpus run observed, grouped so a human can act on it.
 *
 * A run is sharded, so a census must merge: counts add, signatures union, and
 * every field stays a function of the observations alone. Nothing here reads
 * the clock or the filesystem, so a merge of N shard censuses equals the census
 * of one unsharded run over the same files.
 */

import { NOT_A_DOCX_REASONS, type NotADocxReason } from "./corpus-classify";
import { isGatingFailure } from "./corpus-invariants/contract";
import {
  type CorpusFailure,
  type CorpusInvariant,
  distinctBySignature,
  failureSignature,
} from "./corpus-signature";

/**
 * What a run learned about one file.
 *
 * `truncated` is the load-dependent case: a per-file budget ran out, so the
 * invariants after it never ran. Which ones those are depends on how busy the
 * machine was, and the ratchet is exact, so a truncated file's gating evidence
 * is not evidence — including the findings it did produce before the budget
 * ran out, because a slower run would have stopped earlier and reported fewer.
 *
 * `report-only` is the committed case: the file is named in
 * `corpus/report-only-files.json`, so it is measured and reported but never
 * gated, whether it finished or not. Which kind a file gets is decided once,
 * where the run turns an outcome into a result, and the builder reads this
 * with an exhaustive switch, so a gating signature cannot be emitted for
 * either kind by forgetting a check somewhere.
 */
export type CorpusFileResult =
  | { kind: "complete"; failures: readonly CorpusFailure[] }
  | { kind: "truncated"; failures: readonly CorpusFailure[]; stage: string }
  | { kind: "report-only"; failures: readonly CorpusFailure[] }
  | { kind: "not-a-docx"; reason: NotADocxReason };

/** Whether a file's findings may reach a baseline, or only the report. */
export const CORPUS_EVIDENCE = {
  gating: "gating",
  reportOnly: "report-only",
} as const;

export type CorpusEvidence = (typeof CORPUS_EVIDENCE)[keyof typeof CORPUS_EVIDENCE];

/**
 * The one rule that says whether a result's findings gate.
 *
 * Both censuses read it, so the core census and the family censuses cannot
 * disagree about a file, and a result kind added later has to answer here.
 */
export const evidenceOf = (result: CorpusFileResult): CorpusEvidence => {
  switch (result.kind) {
    case "complete": {
      return CORPUS_EVIDENCE.gating;
    }
    case "truncated":
    case "report-only":
    case "not-a-docx": {
      return CORPUS_EVIDENCE.reportOnly;
    }
    default: {
      return assertNever(result);
    }
  }
};

/** A file whose run stopped early, and the stage it stopped at. */
export type TruncatedFile = { file: CorpusFileId; stage: string };

/** A file identified the way a reader can find it again: source, path, content. */
export type CorpusFileId = {
  sourceId: string;
  path: string;
  sha256: string;
};

/**
 * How every part of the gate names a file.
 *
 * One implementation: a selection, a census line and the report-only list all
 * match on this string, so two spellings of it would silently stop agreeing
 * about which file is which.
 */
export const fileIdOf = ({ sourceId, path }: { sourceId: string; path: string }): string =>
  `${sourceId}/${path}`;

export type CensusSignature = {
  signature: string;
  invariant: CorpusInvariant;
  message: string;
  frame: string;
  files: number;
  examples: CorpusFileId[];
};

export type CorpusCensus = {
  schemaVersion: 1;
  lockDigest: string;
  /** The `corpus/report-only-files.json` this run read, so a baseline can bind to it. */
  reportOnlyDigest: string;
  files: number;
  duplicates: number;
  notADocx: number;
  notADocxByReason: Record<NotADocxReason, number>;
  /** Up to three files per reason, so a classification can be audited without rerunning. */
  notADocxExamples: Record<NotADocxReason, CorpusFileId[]>;
  passed: number;
  failedFiles: number;
  /** Files whose run stopped at a budget, so they carry no gating evidence. */
  truncated: number;
  truncatedExamples: TruncatedFile[];
  /** Files the committed list excludes from gating evidence, finished or not. */
  reportOnly: number;
  signatures: CensusSignature[];
};

export const MAX_EXAMPLES_PER_SIGNATURE = 3;

const NOT_A_DOCX_REASON_LIST = Object.values(NOT_A_DOCX_REASONS);

const emptyReasonRecord = <T>(value: () => T): Record<NotADocxReason, T> => {
  const record: Partial<Record<NotADocxReason, T>> = {};
  for (const reason of NOT_A_DOCX_REASON_LIST) {
    record[reason] = value();
  }
  // Every member of the union was just assigned, so the record is total.
  return record as Record<NotADocxReason, T>;
};

export const emptyCensus = (lockDigest: string, reportOnlyDigest: string): CorpusCensus => ({
  schemaVersion: 1,
  lockDigest,
  reportOnlyDigest,
  files: 0,
  duplicates: 0,
  notADocx: 0,
  notADocxByReason: emptyReasonRecord(() => 0),
  notADocxExamples: emptyReasonRecord<CorpusFileId[]>(() => []),
  passed: 0,
  failedFiles: 0,
  truncated: 0,
  truncatedExamples: [],
  reportOnly: 0,
  signatures: [],
});

/** Up to this many truncated files are named, so a degraded run can be diagnosed. */
export const MAX_TRUNCATED_EXAMPLES = 25;

const isReportOnly = (failure: CorpusFailure): boolean => !isGatingFailure(failure);

const assertNever = (value: never): never => {
  throw new Error(`unhandled corpus file result: ${JSON.stringify(value)}`);
};

const compareSignatures = (left: CensusSignature, right: CensusSignature): number => {
  if (left.files !== right.files) {
    return right.files - left.files;
  }
  return left.signature < right.signature ? -1 : 1;
};

export class CensusBuilder {
  readonly #census: CorpusCensus;
  readonly #bySignature = new Map<string, CensusSignature>();

  constructor(lockDigest: string, reportOnlyDigest: string) {
    this.#census = emptyCensus(lockDigest, reportOnlyDigest);
  }

  countDuplicate(): void {
    this.#census.duplicates += 1;
  }

  addNotADocx(file: CorpusFileId, reason: NotADocxReason): void {
    this.#census.files += 1;
    this.#census.notADocx += 1;
    this.#census.notADocxByReason[reason] += 1;
    const examples = this.#census.notADocxExamples[reason];
    if (examples.length < MAX_EXAMPLES_PER_SIGNATURE) {
      examples.push(file);
    }
  }

  /**
   * Record one file's result.
   *
   * The switch is exhaustive, so a result kind added later has to say here
   * whether its findings gate.
   */
  add(file: CorpusFileId, result: CorpusFileResult): void {
    switch (result.kind) {
      case "not-a-docx": {
        this.addNotADocx(file, result.reason);
        return;
      }
      case "truncated": {
        this.#census.files += 1;
        this.#census.truncated += 1;
        if (this.#census.truncatedExamples.length < MAX_TRUNCATED_EXAMPLES) {
          this.#census.truncatedExamples.push({ file, stage: result.stage });
        }
        // Only the report-only findings survive: the gating ones this file did
        // produce are as load-dependent as the ones it never reached.
        this.addChecked(file, result.failures.filter(isReportOnly), { counted: false });
        return;
      }
      case "report-only": {
        this.#census.files += 1;
        this.#census.reportOnly += 1;
        // Listed by hand, so this holds however the run went: the file is
        // measured and reported, and neither passes nor fails.
        this.addChecked(file, result.failures.filter(isReportOnly), { counted: false });
        return;
      }
      case "complete": {
        this.addChecked(file, result.failures);
        return;
      }
      default: {
        return assertNever(result);
      }
    }
  }

  addChecked(
    file: CorpusFileId,
    failures: readonly CorpusFailure[],
    { counted = true }: { counted?: boolean } = {},
  ): void {
    if (counted) {
      this.#census.files += 1;
    }
    if (failures.length === 0) {
      if (counted) {
        this.#census.passed += 1;
      }
      return;
    }
    if (counted) {
      this.#census.failedFiles += 1;
    }
    for (const failure of distinctBySignature(failures)) {
      const signature = failureSignature(failure);
      const existing = this.#bySignature.get(signature);
      if (existing === undefined) {
        this.#bySignature.set(signature, {
          signature,
          invariant: failure.invariant,
          message: failure.message,
          frame: failure.frame,
          files: 1,
          examples: [file],
        });
        continue;
      }
      existing.files += 1;
      if (existing.examples.length < MAX_EXAMPLES_PER_SIGNATURE) {
        existing.examples.push(file);
      }
    }
  }

  build(): CorpusCensus {
    return {
      ...this.#census,
      signatures: [...this.#bySignature.values()].sort(compareSignatures),
    };
  }
}

export const mergeCensuses = (censuses: readonly CorpusCensus[]): CorpusCensus => {
  const first = censuses.at(0);
  if (first === undefined) {
    throw new Error("A census merge needs at least one census");
  }
  const merged = emptyCensus(first.lockDigest, first.reportOnlyDigest);
  const bySignature = new Map<string, CensusSignature>();
  for (const census of censuses) {
    merged.files += census.files;
    merged.duplicates += census.duplicates;
    merged.notADocx += census.notADocx;
    merged.passed += census.passed;
    merged.failedFiles += census.failedFiles;
    merged.truncated += census.truncated ?? 0;
    merged.reportOnly += census.reportOnly ?? 0;
    merged.truncatedExamples.push(
      ...(census.truncatedExamples ?? []).slice(
        0,
        MAX_TRUNCATED_EXAMPLES - merged.truncatedExamples.length,
      ),
    );
    for (const reason of NOT_A_DOCX_REASON_LIST) {
      merged.notADocxByReason[reason] += census.notADocxByReason[reason] ?? 0;
      const examples = merged.notADocxExamples[reason];
      examples.push(
        ...(census.notADocxExamples[reason] ?? []).slice(
          0,
          MAX_EXAMPLES_PER_SIGNATURE - examples.length,
        ),
      );
    }
    // A census written by a build that knows a reason this one does not would
    // lose those files silently; refusing is the only honest merge.
    const unknown = Object.keys(census.notADocxByReason).filter(
      (reason) => !(reason in merged.notADocxByReason),
    );
    if (unknown.length > 0) {
      throw new Error(`Unknown not-a-docx reasons in census: ${unknown.join(", ")}`);
    }
    for (const signature of census.signatures) {
      const existing = bySignature.get(signature.signature);
      if (existing === undefined) {
        bySignature.set(signature.signature, { ...signature, examples: [...signature.examples] });
        continue;
      }
      existing.files += signature.files;
      existing.examples.push(
        ...signature.examples.slice(0, MAX_EXAMPLES_PER_SIGNATURE - existing.examples.length),
      );
    }
  }
  merged.signatures = [...bySignature.values()].sort(compareSignatures);
  return merged;
};

export const renderCensus = (census: CorpusCensus, topSignatures: number): string => {
  const lines = [
    `files ${census.files} (${census.duplicates} duplicate content skipped)`,
    `  not-a-docx ${census.notADocx}: ${NOT_A_DOCX_REASON_LIST.filter(
      (reason) => census.notADocxByReason[reason] > 0,
    )
      .map((reason) => `${reason} ${census.notADocxByReason[reason]}`)
      .join(", ")}`,
    `  passed ${census.passed}, failed ${census.failedFiles}, truncated ${census.truncated}, report-only ${census.reportOnly}, signatures ${census.signatures.length}`,
  ];
  for (const signature of census.signatures.slice(0, topSignatures)) {
    const example = signature.examples.at(0);
    lines.push(
      `  ${String(signature.files).padStart(5)}  ${signature.signature}`,
      `         e.g. ${example === undefined ? "-" : `${example.sourceId}/${example.path}`}`,
    );
  }
  return lines.join("\n");
};
