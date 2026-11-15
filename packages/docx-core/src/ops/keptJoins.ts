/**
 * The joins a pending tracked change takes over from a resolved piece that
 * meets it.
 */

import { panic } from "better-result";

import type { InlineNode } from "./leaves";
import { isTrackedWrapper } from "./review";

/** The range boundaries a removal kept standing at its seam. */
export type KeptAtSeam = {
  /** The date of the revision whose content they were kept from. */
  from: string;
  /** How many of the boundaries standing at the seam they are. */
  count: number;
};

/**
 * `node`, a pending tracked change, with its join on `side` at least `depth`
 * levels deep. A resolution keeping a piece of a record whose join meets a
 * pending insertion made inside that record hands the join to the
 * insertion: accepting the insertion later joins its content with the piece
 * as deep as accepting both at once does. A rejection removing content made
 * inside the record hands its seam on the same way, with `kept`, the range
 * boundaries that content kept standing at the seam: the later join places
 * them inside the record it makes. An insertion whose join on that side is
 * already as deep holds a piece of that record, made inside it, and takes
 * the boundaries all the same, so accepting it places them inside its piece.
 */
export const withJoinAtLeast = (
  node: InlineNode,
  side: "before" | "after",
  depth: number,
  kept?: KeptAtSeam,
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
  const keptAt =
    kept === undefined || joins.keptFrom !== undefined
      ? undefined
      : { keptFrom: kept.from, keptCount: kept.count };
  if (joins[side] >= depth)
    return keptAt === undefined
      ? node
      : Object.assign({}, node, {
          resolutionJoins: Object.assign({}, joins, keptAt),
        });
  return Object.assign({}, node, {
    resolutionJoins: Object.assign(
      {},
      joins,
      side === "before" ? { before: depth } : { after: depth },
      keptAt ?? {},
    ),
  });
};
