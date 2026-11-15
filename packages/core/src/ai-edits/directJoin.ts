/**
 * Which paragraph a direct join of two paragraphs leaves.
 *
 * A join removes the earlier paragraph's mark, and the joined paragraph ends
 * with the later one's. Accepting the same join tracked keeps the paragraph
 * whose mark remains, so a direct `mergeBlockWithNext`, and a direct
 * `deleteBlock` of a container's final paragraph that a paragraph precedes,
 * leave the later paragraph's identity too: both engines read this one rule.
 */

import { SPLIT_HALVES, type SplitHalf } from "@stll/docx-core/ops";

/** The half of a direct join whose identity the joined paragraph keeps. */
export const directJoinSurvivor = (): SplitHalf => SPLIT_HALVES.SECOND;

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
