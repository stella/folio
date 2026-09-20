/**
 * Removing inline items from a paragraph, without leaving it in a shape the
 * parse would not produce.
 *
 * Its own module rather than a member of `paragraphTraversal`: restoring the
 * consolidation needs `runConsolidator`, which reaches `paragraphTraversal`
 * again through `paragraphPropertySource`.
 */

import type { ParagraphContent } from "../types/document";
import type { InlineContentSlot } from "./paragraphTraversal";
import { consolidateParagraphContent } from "./runConsolidator";

/**
 * Positions marked for removal, per inline-content array.
 *
 * Removal shifts every later index in that array, so a normaliser records the
 * positions while it reads and drops them once, after it has finished reading.
 */
export class InlineContentRemovals {
  readonly #byContent = new Map<ParagraphContent[], Set<number>>();

  mark({ content, index }: Pick<InlineContentSlot, "content" | "index">): void {
    const indexes = this.#byContent.get(content);
    if (indexes) {
      indexes.add(index);
      return;
    }
    this.#byContent.set(content, new Set([index]));
  }

  /**
   * Applies every marked removal and answers how many items were dropped.
   *
   * Consolidation is re-run over each array a removal touched, because every
   * inline item that is not a run is a merge boundary: `parseParagraph`
   * consolidated the array while the removed item still stood between its
   * neighbours, so dropping it leaves two runs adjacent that the next parse
   * merges. That is an oscillation rather than a loss, save 1 writes the pair
   * and save 2 writes one run, and it reaches every normaliser that removes an
   * item, not only the one it was found through. Restoring the invariant here
   * makes the parse a fixed point by construction; an array nothing was
   * removed from is not touched.
   */
  apply(): number {
    let removed = 0;
    for (const [content, indexes] of this.#byContent) {
      if (indexes.size === 0) {
        continue;
      }
      const kept = content.filter((_, index) => !indexes.has(index));
      // Counted before consolidation: merging two runs also shortens the
      // array, and what the caller reports is how many items it removed.
      removed += content.length - kept.length;
      const next = consolidateParagraphContent(kept);
      content.length = 0;
      content.push(...next);
    }
    return removed;
  }
}
