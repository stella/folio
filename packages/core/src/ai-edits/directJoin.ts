/**
 * Which paragraph a direct join of two paragraphs leaves.
 *
 * The joined paragraph is the one whose mark remains, as accepting the same
 * join tracked keeps it. A direct `mergeBlockWithNext`, and a direct
 * `deleteBlock` of a container's final paragraph that a paragraph precedes,
 * remove the earlier paragraph's mark and end the joined paragraph with the
 * later one's: they leave the later paragraph's identity. A paragraph deleted
 * directly before a table or a content control, with a paragraph before it,
 * takes its own mark with it: no later mark remains, and the earlier
 * paragraph keeps its identity. Both engines read this one rule.
 */

import { SPLIT_HALVES, type SplitHalf } from "@stll/docx-core/ops";

/**
 * The half of a direct join whose identity the joined paragraph keeps: the
 * later one, unless the later paragraph's mark goes with it.
 */
export const directJoinSurvivor = (
  { laterMarkRemains }: { laterMarkRemains: boolean } = {
    laterMarkRemains: true,
  },
): SplitHalf => (laterMarkRemains ? SPLIT_HALVES.SECOND : SPLIT_HALVES.FIRST);

/**
 * The paragraphs of a direct join that an operation of the same batch can no
 * longer name once the join has run: the one whose id it retires, and the
 * one it keeps, whose text and offsets it rewrites. Whichever survives, a
 * batch runs every operation that names either before the join, as the same
 * operations one at a time in the batch's order do.
 */
export const directJoinTakes = <Block>(
  first: Block,
  second: Block,
): readonly Block[] => [first, second];
