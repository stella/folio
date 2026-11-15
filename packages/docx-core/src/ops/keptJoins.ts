/**
 * The joins a pending tracked change takes over from a resolved piece that
 * meets it.
 */

import { panic } from "better-result";

import type { InlineNode } from "./leaves";
import { isTrackedWrapper } from "./review";

/**
 * `node`, a pending tracked change, with its join on `side` at least `depth`
 * levels deep. A resolution keeping a piece of a record whose join meets a
 * pending insertion made inside that record hands the join to the
 * insertion: accepting the insertion later joins its content with the piece
 * as deep as accepting both at once does. A rejection removing content made
 * inside the record hands its seam on the same way, with `keptFrom`, the
 * date of that content's revision, when range boundaries it kept stand at
 * the seam: the later join places them inside the record it makes.
 */
export const withJoinAtLeast = (
  node: InlineNode,
  side: "before" | "after",
  depth: number,
  keptFrom?: string,
): InlineNode => {
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
    return node;
  // A join already as deep still takes the date of the boundaries the seam
  // kept, so the later join places them inside the record it makes.
  if (joins[side] >= depth)
    return keptFrom === undefined || joins.keptFrom !== undefined
      ? node
      : Object.assign({}, node, {
          resolutionJoins: Object.assign({}, joins, { keptFrom }),
        });
  return Object.assign({}, node, {
    resolutionJoins: Object.assign(
      {},
      joins,
      side === "before" ? { before: depth } : { after: depth },
      keptFrom === undefined || joins.keptFrom !== undefined
        ? {}
        : { keptFrom },
    ),
  });
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
