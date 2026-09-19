/**
 * The census: what a corpus run observed, grouped so a human can act on it.
 *
 * A run is sharded, so a census must merge: counts add, signatures union, and
 * every field stays a function of the observations alone. Nothing here reads
 * the clock or the filesystem, so a merge of N shard censuses equals the census
 * of one unsharded run over the same files.
 */

import { NOT_A_DOCX_REASONS, type NotADocxReason } from "./corpus-classify";
import { type CorpusFailure, type CorpusInvariant, failureSignature } from "./corpus-signature";

/** A file identified the way a reader can find it again: source, path, content. */
export type CorpusFileId = {
  sourceId: string;
  path: string;
  sha256: string;
};

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
  files: number;
  duplicates: number;
  notADocx: number;
  notADocxByReason: Record<NotADocxReason, number>;
  /** Up to three files per reason, so a classification can be audited without rerunning. */
  notADocxExamples: Record<NotADocxReason, CorpusFileId[]>;
  passed: number;
  failedFiles: number;
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

export const emptyCensus = (lockDigest: string): CorpusCensus => ({
  schemaVersion: 1,
  lockDigest,
  files: 0,
  duplicates: 0,
  notADocx: 0,
  notADocxByReason: emptyReasonRecord(() => 0),
  notADocxExamples: emptyReasonRecord<CorpusFileId[]>(() => []),
  passed: 0,
  failedFiles: 0,
  signatures: [],
});

const compareSignatures = (left: CensusSignature, right: CensusSignature): number => {
  if (left.files !== right.files) {
    return right.files - left.files;
  }
  return left.signature < right.signature ? -1 : 1;
};

export class CensusBuilder {
  readonly #census: CorpusCensus;
  readonly #bySignature = new Map<string, CensusSignature>();

  constructor(lockDigest: string) {
    this.#census = emptyCensus(lockDigest);
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

  addChecked(file: CorpusFileId, failures: readonly CorpusFailure[]): void {
    this.#census.files += 1;
    if (failures.length === 0) {
      this.#census.passed += 1;
      return;
    }
    this.#census.failedFiles += 1;
    for (const failure of failures) {
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
  const merged = emptyCensus(first.lockDigest);
  const bySignature = new Map<string, CensusSignature>();
  for (const census of censuses) {
    merged.files += census.files;
    merged.duplicates += census.duplicates;
    merged.notADocx += census.notADocx;
    merged.passed += census.passed;
    merged.failedFiles += census.failedFiles;
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
    `  passed ${census.passed}, failed ${census.failedFiles}, signatures ${census.signatures.length}`,
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
