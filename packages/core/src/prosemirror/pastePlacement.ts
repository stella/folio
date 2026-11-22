/** Where a paste or a drop puts its content. */
import type { Transaction } from "prosemirror-state";
import { Mapping, type StepMap } from "prosemirror-transform";

/** Positions in a document, end exclusive. */
export type PlacedRange = { from: number; to: number };

export const mapPlacedRanges = (
  ranges: readonly PlacedRange[],
  maps: readonly StepMap[],
): PlacedRange[] =>
  ranges.flatMap(({ from, to }) => {
    const mapping = new Mapping(maps);
    // oxlint-disable-next-line unicorn/no-array-method-this-argument -- ProseMirror Mapping.map(pos, assoc)
    const mapped = { from: mapping.map(from, 1), to: mapping.map(to, -1) };
    return mapped.from < mapped.to ? [mapped] : [];
  });

/** One place holding content its paste put in an earlier place already. */
export type PastedRepeat = readonly PlacedRange[];

const PASTED_REPEATS_META = "folioPastedRepeats";

/**
 * Record where a paste put the same content again, as a paste into several
 * selected table cells does, in the transaction's document.
 */
export const setPastedRepeats = (
  transaction: Transaction,
  repeats: readonly PastedRepeat[],
): void => {
  if (repeats.length > 0) transaction.setMeta(PASTED_REPEATS_META, repeats);
};

export const pastedRepeatsOf = (
  transaction: Transaction,
): readonly PastedRepeat[] => {
  // SAFETY: `setPastedRepeats` is the only writer of this meta.
  const repeats = transaction.getMeta(PASTED_REPEATS_META) as
    readonly PastedRepeat[] | undefined;
  return repeats ?? [];
};
