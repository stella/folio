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
    joins.plainSides?.includes(side) === true ||
    joins[side] >= depth
  )
    return node;
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
 * Whether a seam this resolution records hands its join to the pending
 * insertions facing it. A kept revision's join always does: the insertion
 * was made inside the record the join belongs to. A removal's seam belongs
 * to the pieces of the record its content cut. With pieces standing beyond
 * the pending insertions on both sides, the insertions between are new
 * content, no piece of that record: the seam is deferred on them, so
 * rejecting them joins the pieces and accepting them keeps the pieces apart,
 * as resolving both at once does. Only without pieces beyond both are the
 * insertions facing the seam its pieces, and they take its join.
 */
export const seamHandsOver = ({
  seam,
  piecesBeyond,
}: {
  seam: "keptJoin" | "removal";
  piecesBeyond: boolean;
}): boolean => seam === "keptJoin" || !piecesBeyond;
