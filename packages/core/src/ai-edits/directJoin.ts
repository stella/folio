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
