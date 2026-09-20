/**
 * What changed between two corpus runs, per family and per file.
 *
 * A baseline comparison answers "is this worse than what we recorded". A
 * differential answers "what did this commit do", which the baseline cannot:
 * it carries counts, and a count that moved from 30 to 137 names neither the
 * files that arrived nor the ones that left. Reconstructing them meant running
 * both revisions again with `--only` over guesses, so nobody did.
 *
 * This reads the per-file rows two runs wrote and reports, per family, the
 * signatures only the later run has, the ones only the earlier run has, and
 * the ones both have whose file set moved. A signature's growth is reported as
 * the files that reached it, by id, which is the question a differential is
 * always asked next.
 */

import type { CorpusInvariantFamily } from "./corpus-invariants/contract";
import {
  type CorpusFileSignatures,
  type FilesBySignature,
  indexFileSignatures,
} from "./corpus-file-signatures";

/** Why a signature appears in the report. */
export const SIGNATURE_DELTA_KINDS = {
  /** Only the later run has it: the commit created this failure mode. */
  introduced: "INTRODUCED",
  /** Only the earlier run has it: nothing fails this way any more. */
  fixed: "FIXED",
  /** Both have it, and the set of files it fires on moved. */
  reachedNow: "REACHED-NOW",
} as const;

export type SignatureDeltaKind = (typeof SIGNATURE_DELTA_KINDS)[keyof typeof SIGNATURE_DELTA_KINDS];

export type SignatureDelta = {
  kind: SignatureDeltaKind;
  signature: string;
  before: number;
  after: number;
  /** Files that fail this way now and did not before, sorted. */
  reached: string[];
  /** Files that failed this way before and no longer do, sorted. */
  left: string[];
};

export type FamilyDelta = {
  family: CorpusInvariantFamily;
  deltas: SignatureDelta[];
};

const sorted = <T extends string>(ids: Iterable<T>): T[] => [...ids].sort();

const difference = (left: ReadonlySet<string>, right: ReadonlySet<string>): string[] =>
  sorted([...left].filter((id) => !right.has(id)));

const EMPTY: ReadonlySet<string> = new Set();

const deltaKind = (before: number, after: number): SignatureDeltaKind => {
  if (before === 0) {
    return SIGNATURE_DELTA_KINDS.introduced;
  }
  return after === 0 ? SIGNATURE_DELTA_KINDS.fixed : SIGNATURE_DELTA_KINDS.reachedNow;
};

/** Biggest movement first, then by signature so a diff is byte-stable. */
const compareDeltas = (left: SignatureDelta, right: SignatureDelta): number => {
  const moved = right.reached.length + right.left.length - (left.reached.length + left.left.length);
  if (moved !== 0) {
    return moved;
  }
  return left.signature < right.signature ? -1 : 1;
};

const deltasFor = (before: FilesBySignature, after: FilesBySignature): SignatureDelta[] => {
  const deltas: SignatureDelta[] = [];
  for (const signature of new Set([...before.keys(), ...after.keys()])) {
    const was = before.get(signature) ?? EMPTY;
    const now = after.get(signature) ?? EMPTY;
    const reached = difference(now, was);
    const left = difference(was, now);
    if (reached.length === 0 && left.length === 0) {
      continue;
    }
    deltas.push({
      kind: deltaKind(was.size, now.size),
      signature,
      before: was.size,
      after: now.size,
      reached,
      left,
    });
  }
  return deltas.sort(compareDeltas);
};

export const diffFileSignatures = (
  before: readonly CorpusFileSignatures[],
  after: readonly CorpusFileSignatures[],
): FamilyDelta[] => {
  const indexedBefore = indexFileSignatures(before);
  const indexedAfter = indexFileSignatures(after);
  const families = sorted(new Set([...indexedBefore.keys(), ...indexedAfter.keys()]));
  const empty: FilesBySignature = new Map();
  return families
    .map((family) => ({
      family,
      deltas: deltasFor(indexedBefore.get(family) ?? empty, indexedAfter.get(family) ?? empty),
    }))
    .filter(({ deltas }) => deltas.length > 0);
};

/** Files named per signature; past this the report says how many more. */
export const MAX_NAMED_FILES = 40;

const namedFiles = (label: string, ids: readonly string[]): string[] => {
  if (ids.length === 0) {
    return [];
  }
  const shown = ids.slice(0, MAX_NAMED_FILES);
  const rest = ids.length - shown.length;
  return [
    `      ${label} ${ids.length}${rest > 0 ? ` (${rest} more not named)` : ""}:`,
    ...shown.map((id) => `        ${id}`),
  ];
};

export const renderFamilyDeltas = (families: readonly FamilyDelta[]): string => {
  if (families.length === 0) {
    return "No signature changed its file set.";
  }
  const lines: string[] = [];
  for (const { family, deltas } of families) {
    lines.push(`${family}: ${deltas.length} signature(s) moved`);
    for (const { kind, signature, before, after, reached, left } of deltas) {
      lines.push(`  [${kind}] ${before} -> ${after} files`, `    ${signature}`);
      lines.push(...namedFiles("+", reached), ...namedFiles("-", left));
    }
  }
  return lines.join("\n");
};
