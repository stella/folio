/**
 * Block lookup by id for passes over a layout's fragments.
 */

import type { BlockId, FlowBlock } from "../../layout-engine/types";

/**
 * Each block's index by id, the first block's where ids repeat. A pass that
 * looks up the block of every fragment builds it once instead of searching
 * the blocks for each fragment.
 */
export const indexBlocksById = (
  blocks: readonly FlowBlock[],
): ReadonlyMap<BlockId, number> => {
  const indices = new Map<BlockId, number>();
  for (const [index, block] of blocks.entries()) {
    if (!indices.has(block.id)) indices.set(block.id, index);
  }
  return indices;
};
