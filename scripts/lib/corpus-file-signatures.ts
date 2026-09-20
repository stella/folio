/**
 * Which files a signature fired on, all of them.
 *
 * The census keeps three examples per signature, which is the right size for a
 * report and useless for a differential. When a row grew by a hundred files
 * between two commits, "e.g. one of them" does not say which hundred, so the
 * only way to attribute the growth was to rerun both sides with `--only` over
 * guesses. This records the whole mapping instead: one row per file per
 * invariant family, holding the signatures that family counted for it.
 *
 * It is the census's own verdict, not a second opinion: the signatures come
 * from `countedFailures`, which is the rule the family census counts by, so an
 * attribution can never name a signature the census did not count or miss one
 * it did.
 *
 * The rows are large, so a run only produces them when asked. Nothing ratchets
 * against them and nothing else reads them; the census is still the artifact
 * the gate compares.
 */

import type { CorpusFileId } from "./corpus-census";
import { fileIdOf } from "./corpus-census";
import { type CorpusInvariantFamily, familyOf } from "./corpus-invariants/contract";
import type { CorpusFailure } from "./corpus-signature";
import { failureSignature } from "./corpus-signature";

export type CorpusFileSignatures = {
  /** `sourceId/path`: the id the baselines, the examples and `--only` all use. */
  id: string;
  family: CorpusInvariantFamily;
  /** Every signature this family counted for this file, sorted. */
  signatures: string[];
};

/**
 * One file's rows, one per family it has findings in.
 *
 * A file with no findings produces nothing: the diff asks which files a
 * signature reached, and a file that reached none is the ordinary case.
 */
export const fileSignatureRows = (
  file: CorpusFileId,
  counted: readonly CorpusFailure[],
): CorpusFileSignatures[] => {
  const byFamily = new Map<CorpusInvariantFamily, string[]>();
  for (const failure of counted) {
    const family = familyOf(failure.invariant);
    const signatures = byFamily.get(family) ?? [];
    signatures.push(failureSignature(failure));
    byFamily.set(family, signatures);
  }
  const id = fileIdOf(file);
  return [...byFamily.entries()]
    .map(([family, signatures]) => ({ id, family, signatures: signatures.sort() }))
    .sort((left, right) => (left.family < right.family ? -1 : 1));
};

/** Signature to the ids of every file it fired on, for one family. */
export type FilesBySignature = Map<string, Set<string>>;

export type SignaturesByFamily = Map<CorpusInvariantFamily, FilesBySignature>;

/**
 * The rows inverted, which is the shape a differential reads.
 *
 * Shards write disjoint rows, so several artifacts index into one map without
 * a merge step of their own.
 */
export const indexFileSignatures = (rows: readonly CorpusFileSignatures[]): SignaturesByFamily => {
  const byFamily: SignaturesByFamily = new Map();
  for (const { id, family, signatures } of rows) {
    const bySignature = byFamily.get(family) ?? new Map<string, Set<string>>();
    for (const signature of signatures) {
      const files = bySignature.get(signature) ?? new Set<string>();
      files.add(id);
      bySignature.set(signature, files);
    }
    byFamily.set(family, bySignature);
  }
  return byFamily;
};
