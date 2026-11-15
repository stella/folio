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
 * as deep as accepting both at once does.
 */
export const withJoinAtLeast = (
  node: InlineNode,
  side: "before" | "after",
  depth: number,
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
    ),
  });
};
