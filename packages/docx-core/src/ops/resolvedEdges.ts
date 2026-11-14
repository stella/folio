/**
 * The records at an edge of an inline record, as a resolution reads them
 * when it measures how far a cut reaches.
 */

import { childNodes, type InlineNode } from "./leaves";

/** `node` and its last (`end`) or first (`start`) child, theirs, and so on. */
export const edgeChain = (
  node: InlineNode,
  edge: "start" | "end",
): InlineNode[] => {
  const chain: InlineNode[] = [];
  let current: InlineNode | undefined = node;
  while (current !== undefined) {
    chain.push(current);
    const children = childNodes(current);
    current = edge === "start" ? children?.at(0) : children?.at(-1);
  }
  return chain;
};

/**
 * The edge chain `chain` becomes once its record at `index`, which the
 * resolution removes with its content, is gone: what then stands at that
 * edge of the record above it is the next record inward, so the chain goes
 * through that one from `index` on. Read so, a cut reaches the record the
 * removed one stood in front of, as it does when the removal is resolved on
 * its own first. `undefined` when the removed record has no record above it
 * in the chain or nothing else stands beside it in that record.
 */
export const edgeChainPastRemoved = (
  chain: readonly InlineNode[],
  index: number,
  edge: "start" | "end",
): InlineNode[] | undefined => {
  const parent = index > 0 ? chain.at(index - 1) : undefined;
  const removed = chain.at(index);
  if (parent === undefined || removed === undefined) return undefined;
  const siblings = childNodes(parent) ?? [];
  const inward = edge === "start" ? siblings : siblings.toReversed();
  const next = inward.at(inward.indexOf(removed) + 1);
  if (next === undefined || !inward.includes(removed)) return undefined;
  return [...chain.slice(0, index), ...edgeChain(next, edge)];
};
