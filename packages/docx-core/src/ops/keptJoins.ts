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
  if (node.resolutionJoins[side] >= depth) return node;
  return Object.assign({}, node, {
    resolutionJoins: Object.assign(
      {},
      node.resolutionJoins,
      side === "before" ? { before: depth } : { after: depth },
    ),
  });
};
