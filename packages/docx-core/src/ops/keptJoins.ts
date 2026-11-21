/**
 * The joins a pending tracked change takes over from a resolved piece that
 * meets it.
 */

import { panic } from "better-result";

import type { InlineNode } from "./leaves";
import { isTrackedWrapper } from "./review";

/**
 * A pending tracked change handed a seam's join (see {@link handJoin}), and
 * whether it takes the kept range markers standing at that seam.
 */
export type HandedJoin = { node: InlineNode; takesKept: boolean };

/**
 * `node`, a pending tracked change, with its join on `side` at least `depth`
 * levels deep. A resolution keeping a piece of a record whose join meets a
 * pending insertion made inside that record hands the join to the
 * insertion: accepting the insertion later joins its content with the piece
 * as deep as accepting both at once does. A rejection removing content made
 * inside the record hands its seam on the same way, with the range markers
 * that content kept standing at the seam: they name the insertion's side as
 * their cut's piece, so the later join places them inside the record it
 * makes. An insertion whose join on that side is already as deep holds a
 * piece of that record, made inside it, and takes the markers all the same,
 * so accepting it places them inside its piece.
 */
export const handJoin = (
  node: InlineNode,
  side: "before" | "after",
  depth: number,
): HandedJoin => {
  if (!isTrackedWrapper(node) || node.resolutionJoins === undefined)
    return panic("A blocking insertion is a pending change with joins.");
  const joins = node.resolutionJoins;
  // Inserted text whose acceptance joins only plain runs, or that meets
  // accepted new content on that side, is no piece of the record the join
  // belongs to and takes none.
  if (
    joins.acceptance !== undefined ||
    joins.plainSides?.includes(side) === true
  )
    return { node, takesKept: false };
  if (joins[side] >= depth) return { node, takesKept: true };
  return {
    node: Object.assign({}, node, {
      resolutionJoins: Object.assign(
        {},
        joins,
        side === "before" ? { before: depth } : { after: depth },
      ),
    }),
    takesKept: true,
  };
};

/**
 * Whether a seam this resolution records hands its join to a pending
 * insertion facing it, whose own join on that side is `facingJoin`. A kept
 * revision's join always does: the insertion was made inside the record the
 * join belongs to. A removal's seam belongs to the pieces of the record its
 * content cut. An insertion joined to the seam's side holds such a piece,
 * made inside that record, and takes the seam. One joined to nothing there
 * is new content: with pieces of the record beyond the pending insertions on
 * both sides, the seam is deferred on it, so rejecting it joins the pieces
 * and accepting it keeps them apart, as resolving both at once does. Only
 * without pieces beyond both are the insertions facing the seam its pieces.
 */
export const seamHandsOver = ({
  seam,
  piecesBeyond,
  facingJoin,
}: {
  seam: "keptJoin" | "removal";
  piecesBeyond: boolean;
  facingJoin: number;
}): boolean => seam === "keptJoin" || facingJoin > 0 || !piecesBeyond;

/**
 * `node`, a pending tracked change, with no join on `side`: the content it
 * faced there, made inside its piece, was rejected, so the join has nothing
 * left to take.
 */
export const withJoinRetired = (
  node: InlineNode,
  side: "before" | "after",
): InlineNode => {
  if (!isTrackedWrapper(node) || node.resolutionJoins === undefined)
    return panic("A facing insertion is a pending change with joins.");
  const joins = node.resolutionJoins;
  if (joins[side] === 0) return node;
  return Object.assign({}, node, {
    resolutionJoins: Object.assign(
      {},
      joins,
      side === "before" ? { before: 0 } : { after: 0 },
    ),
  });
};
