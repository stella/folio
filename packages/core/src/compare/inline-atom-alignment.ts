/**
 * How a comparison lines up the inline atoms (fields, pictures, page breaks)
 * of one reviewed block with the revised block it stands for, whatever the
 * atoms are read from. A reader gives each block its clean text, its supported
 * atoms with their clean-text offsets and identity keys, and the zero-width
 * topology it cannot restore; everything here is decided on those alone.
 */

import { canonicalJson } from "../utils/canonicalJson";

/** One supported atom of a block, as alignment sees it. */
export type AlignableAtom = {
  /** Clean-text characters before the atom in its block. */
  readonly offset: number;
  /** The atom's document facts, equal exactly when the atoms are the same. */
  readonly key: string;
  /** The field result the atom shows as text, or `null` for any other atom. */
  readonly resultText: string | null;
};

/** One block of a story, as alignment sees it. */
export type AlignableBlock = {
  /** The block's kind: blocks of two kinds never align. */
  readonly blockType: string;
  readonly cleanText: string;
  readonly supported: readonly AlignableAtom[];
  /** Inline atoms the comparison cannot restore, which must agree where it restores one. */
  readonly unsupportedTopology: readonly {
    readonly offset: number;
    readonly key: string;
  }[];
};

/** A revised reference names a note the redline's package does not have. */
export const MISSING_NOTE_DETAIL =
  "a note reference names a footnote or endnote the base document does not have";

// Inline atom groups normally contain only a few zero-width carriers. Keep the
// exact ordered match bounded so a hostile document cannot make comparison
// quadratic; larger groups retain only stable unique anchors.
export const MAX_INLINE_ATOM_LCS_CELLS = 16_384;

/** Whether two blocks read the same and agree on the atoms beside supported ones. */
export const sameBlockTopology = (
  left: AlignableBlock,
  right: AlignableBlock,
): boolean => {
  if (left.blockType !== right.blockType || left.cleanText !== right.cleanText)
    return false;
  const supportedOffsets = new Set([
    ...left.supported.map(({ offset }) => offset),
    ...right.supported.map(({ offset }) => offset),
  ]);
  const relevantTopology = ({ unsupportedTopology }: AlignableBlock) =>
    unsupportedTopology.filter(({ offset }) => supportedOffsets.has(offset));
  return (
    canonicalJson(relevantTopology(left)) ===
    canonicalJson(relevantTopology(right))
  );
};

const supportedAtomProjection = (block: AlignableBlock) =>
  canonicalJson(block.supported.map(({ offset, key }) => ({ offset, key })));

/** Whether two blocks hold the same supported atoms at the same offsets. */
export const sameSupportedAtoms = (
  left: AlignableBlock,
  right: AlignableBlock,
): boolean =>
  left.blockType === right.blockType &&
  supportedAtomProjection(left) === supportedAtomProjection(right);

/** Which field results each clean-text offset's atoms show. */
export const fieldResultOwnershipOf = (
  atoms: readonly AlignableAtom[],
): ReadonlyMap<number, ReadonlySet<string>> => {
  const ownership = new Map<number, Set<string>>();
  for (const atom of atoms) {
    if (!atom.resultText) continue;
    let resultsAtOffset = ownership.get(atom.offset);
    if (!resultsAtOffset) {
      resultsAtOffset = new Set<string>();
      ownership.set(atom.offset, resultsAtOffset);
    }
    resultsAtOffset.add(atom.resultText);
  }
  return ownership;
};

type OwnsFieldResultOptions = {
  ownership: ReadonlyMap<number, ReadonlySet<string>>;
  offset: number;
  resultText: string;
};

/** Whether the other side has a field showing `resultText` at `offset`. */
export const ownsFieldResult = ({
  ownership,
  offset,
  resultText,
}: OwnsFieldResultOptions): boolean =>
  ownership.get(offset)?.has(resultText) ?? false;

type AtomGroupOptions = {
  live: readonly AlignableAtom[];
  target: readonly AlignableAtom[];
  liveStart: number;
  targetStart: number;
};

const matchingAtomGroup = ({
  live,
  target,
  liveStart,
  targetStart,
}: AtomGroupOptions): readonly [number, number][] => {
  if (
    live.length === target.length &&
    live.every((atom, index) => atom.key === target[index]?.key)
  ) {
    return live.map((_, index) => [liveStart + index, targetStart + index]);
  }
  if (live.length === 0 || target.length === 0) return [];

  if (live.length <= Math.floor(MAX_INLINE_ATOM_LCS_CELLS / target.length)) {
    const width = target.length + 1;
    const table = new Uint16Array((live.length + 1) * width);
    for (let liveIndex = live.length - 1; liveIndex >= 0; liveIndex--) {
      for (
        let targetIndex = target.length - 1;
        targetIndex >= 0;
        targetIndex--
      ) {
        const index = liveIndex * width + targetIndex;
        if (live[liveIndex]?.key === target[targetIndex]?.key) {
          table[index] =
            (table[(liveIndex + 1) * width + targetIndex + 1] ?? 0) + 1;
          continue;
        }
        table[index] = Math.max(
          table[(liveIndex + 1) * width + targetIndex] ?? 0,
          table[liveIndex * width + targetIndex + 1] ?? 0,
        );
      }
    }
    const matches: [number, number][] = [];
    let liveIndex = 0;
    let targetIndex = 0;
    while (liveIndex < live.length && targetIndex < target.length) {
      if (live[liveIndex]?.key === target[targetIndex]?.key) {
        matches.push([liveStart + liveIndex, targetStart + targetIndex]);
        liveIndex++;
        targetIndex++;
        continue;
      }
      const skipLive = table[(liveIndex + 1) * width + targetIndex] ?? 0;
      const skipTarget = table[liveIndex * width + targetIndex + 1] ?? 0;
      if (skipLive >= skipTarget) liveIndex++;
      else targetIndex++;
    }
    return matches;
  }

  const targetIndexes = new Map<string, number | null>();
  for (const [index, atom] of target.entries()) {
    targetIndexes.set(atom.key, targetIndexes.has(atom.key) ? null : index);
  }
  const liveKeyCounts = new Map<string, number>();
  for (const atom of live) {
    liveKeyCounts.set(atom.key, (liveKeyCounts.get(atom.key) ?? 0) + 1);
  }
  const matches: [number, number][] = [];
  let previousTargetIndex = -1;
  for (const [liveIndex, atom] of live.entries()) {
    const targetIndex = targetIndexes.get(atom.key);
    if (
      targetIndex === undefined ||
      targetIndex === null ||
      targetIndex <= previousTargetIndex ||
      liveKeyCounts.get(atom.key) !== 1
    ) {
      continue;
    }
    matches.push([liveStart + liveIndex, targetStart + targetIndex]);
    previousTargetIndex = targetIndex;
  }
  return matches;
};

type MatchingAtomsOptions = {
  live: readonly AlignableAtom[];
  target: readonly AlignableAtom[];
};

/**
 * The atoms both blocks keep, as `[liveIndex, targetIndex]` pairs: atoms at
 * one clean-text offset are matched in order among themselves, never across
 * offsets.
 */
export const matchingAtoms = ({
  live,
  target,
}: MatchingAtomsOptions): readonly [number, number][] => {
  const matched: [number, number][] = [];
  let left = 0;
  let right = 0;
  while (left < live.length || right < target.length) {
    const liveOffset = live[left]?.offset;
    const targetOffset = target[right]?.offset;
    const offset = Math.min(
      liveOffset ?? Number.POSITIVE_INFINITY,
      targetOffset ?? Number.POSITIVE_INFINITY,
    );
    if (!Number.isFinite(offset)) break;
    const liveStart = left;
    const targetStart = right;
    while (live[left]?.offset === offset) left++;
    while (target[right]?.offset === offset) right++;
    matched.push(
      ...matchingAtomGroup({
        live: live.slice(liveStart, left),
        target: target.slice(targetStart, right),
        liveStart,
        targetStart,
      }),
    );
  }
  return matched;
};

/** Where an unmatched revised atom goes: beside a matched reviewed atom of its offset. */
export type InsertionAnchor = {
  readonly liveIndex: number;
  readonly side: "before" | "after";
};

type InsertionAnchorsOptions = MatchingAtomsOptions & {
  matches: readonly (readonly [number, number])[];
};

/**
 * For each revised atom no reviewed atom matches, the matched reviewed atom
 * at the same offset it goes before (the next one) or, failing that, after
 * (the previous one). Atoms of an offset with no match have no anchor.
 */
export const insertionAnchors = ({
  live,
  target,
  matches,
}: InsertionAnchorsOptions): ReadonlyMap<number, InsertionAnchor> => {
  const liveIndexByTargetIndex = new Map<number, number>();
  for (const [liveIndex, targetIndex] of matches) {
    liveIndexByTargetIndex.set(targetIndex, liveIndex);
  }
  const anchors = new Map<number, InsertionAnchor>();
  let groupStart = 0;
  while (groupStart < target.length) {
    const offset = target[groupStart]?.offset;
    let groupEnd = groupStart + 1;
    while (target[groupEnd]?.offset === offset) groupEnd++;
    let nextMatchedLiveIndex: number | undefined;
    for (let index = groupEnd - 1; index >= groupStart; index--) {
      const matchedLiveIndex = liveIndexByTargetIndex.get(index);
      if (matchedLiveIndex !== undefined)
        nextMatchedLiveIndex = matchedLiveIndex;
      else if (nextMatchedLiveIndex !== undefined)
        anchors.set(index, { liveIndex: nextMatchedLiveIndex, side: "before" });
    }
    let previousMatchedLiveIndex: number | undefined;
    for (let index = groupStart; index < groupEnd; index++) {
      const matchedLiveIndex = liveIndexByTargetIndex.get(index);
      if (matchedLiveIndex !== undefined) {
        previousMatchedLiveIndex = matchedLiveIndex;
        continue;
      }
      if (
        !anchors.has(index) &&
        previousMatchedLiveIndex !== undefined &&
        previousMatchedLiveIndex < live.length
      )
        anchors.set(index, {
          liveIndex: previousMatchedLiveIndex,
          side: "after",
        });
    }
    groupStart = groupEnd;
  }
  return anchors;
};

/** What a planned atom action does to content it replaces or removes. */
export type AtomActionDisposition = "inserted" | "retained";

/**
 * Inline presentation ranges an action writes: replacing retained content
 * writes a deletion beside the insertion.
 */
export const atomActionRangeCount = (
  replaces: AtomActionDisposition | null,
): number => (replaces === "retained" ? 2 : 1);
