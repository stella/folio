import type { Selection } from "prosemirror-state";

/**
 * A selection's outer bounds. Selected cells keep the head cell as their
 * primary range, so `from` and `to` alone miss the other cells; the bounds
 * run from the first selected cell to the last, in either direction.
 */
export const selectionBounds = ({
  ranges,
}: Pick<Selection, "ranges">): { from: number; to: number } => {
  let from = Number.POSITIVE_INFINITY;
  let to = Number.NEGATIVE_INFINITY;
  for (const { $from, $to } of ranges) {
    from = Math.min(from, $from.pos);
    to = Math.max(to, $to.pos);
  }
  return Number.isFinite(from) ? { from, to } : { from: 0, to: 0 };
};
